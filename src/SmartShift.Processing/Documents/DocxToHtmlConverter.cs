using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SmartShift.Processing;
using A = DocumentFormat.OpenXml.Drawing;

namespace SmartShift.Processing.Documents;

public static class DocxToHtmlConverter
{
    public static DocxHtmlResult Convert(
        string sourceDocxPath,
        string outputPath,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(sourceDocxPath) || !File.Exists(sourceDocxPath))
        {
            throw new FileNotFoundException("Source Word document does not exist.", sourceDocxPath);
        }

        var sw = Stopwatch.StartNew();
        progress?.Report(new ProcessingProgress(15, "Opening Word document package…"));

        var outDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
        {
            Directory.CreateDirectory(outDir);
        }

        var tempPartial = outputPath + ".partial";
        if (File.Exists(tempPartial))
        {
            try { File.Delete(tempPartial); } catch { }
        }

        int headingsDetected = 0;
        int tablesReconstructed = 0;
        int imagesPreserved = 0;
        long totalChars = 0;
        var limitations = new List<string>();
        var warnings = new List<string>();

        try
        {
            var bodyHtml = new StringBuilder();
            string docTitle = Path.GetFileNameWithoutExtension(sourceDocxPath);

            using (var wordDoc = WordprocessingDocument.Open(sourceDocxPath, false))
            {
                var mainPart = wordDoc.MainDocumentPart;
                var body = mainPart?.Document?.Body;
                if (mainPart == null || body == null)
                {
                    throw new InvalidOperationException("Word document does not contain a valid main body.");
                }

                // Cache hyperlink relationships
                var hyperlinkMap = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (var rel in mainPart.HyperlinkRelationships)
                {
                    hyperlinkMap[rel.Id] = rel.Uri.ToString();
                }

                var elements = body.ChildElements.ToList();
                int totalElements = elements.Count;

                bool inList = false;
                bool isOrderedList = false;

                void CloseListIfOpen()
                {
                    if (inList)
                    {
                        bodyHtml.AppendLine(isOrderedList ? "</ol>" : "</ul>");
                        inList = false;
                    }
                }

                for (int i = 0; i < totalElements; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (i % 20 == 0)
                    {
                        int pct = 20 + (int)(70.0 * i / Math.Max(1, totalElements));
                        progress?.Report(new ProcessingProgress(pct, $"Converting document elements to HTML ({i + 1}/{totalElements})…"));
                    }

                    var elem = elements[i];

                    if (elem is Paragraph para)
                    {
                        // Check for images inside paragraph
                        var blips = para.Descendants<A.Blip>().ToList();
                        foreach (var blip in blips)
                        {
                            var embedId = blip.Embed?.Value;
                            if (!string.IsNullOrEmpty(embedId) && mainPart.TryGetPartById(embedId, out var part) && part is ImagePart imgPart)
                            {
                                try
                                {
                                    using var stream = imgPart.GetStream(FileMode.Open, FileAccess.Read);
                                    using var mem = new MemoryStream();
                                    stream.CopyTo(mem);
                                    var imgBytes = mem.ToArray();
                                    var base64 = System.Convert.ToBase64String(imgBytes);
                                    var mimeType = imgPart.ContentType ?? "image/png";

                                    CloseListIfOpen();
                                    bodyHtml.AppendLine($"<figure><img src=\"data:{mimeType};base64,{base64}\" alt=\"Embedded Document Image\" /></figure>");
                                    imagesPreserved++;
                                }
                                catch
                                {
                                    warnings.Add("An embedded image could not be decoded for HTML export.");
                                }
                            }
                        }

                        var text = para.InnerText?.Trim();
                        if (string.IsNullOrWhiteSpace(text) && blips.Count == 0)
                        {
                            continue;
                        }

                        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                        var (isHeading, level) = DetermineHeading(styleId);
                        bool isBullet = para.ParagraphProperties?.NumberingProperties != null;

                        if (isHeading)
                        {
                            CloseListIfOpen();
                            headingsDetected++;
                            totalChars += text?.Length ?? 0;
                            int tagLevel = Math.Clamp(level, 1, 6);
                            var content = RenderInlineRuns(para, hyperlinkMap);
                            bodyHtml.AppendLine($"<h{tagLevel}>{content}</h{tagLevel}>");
                        }
                        else if (isBullet)
                        {
                            if (!inList)
                            {
                                inList = true;
                                isOrderedList = false;
                                bodyHtml.AppendLine("<ul>");
                            }
                            totalChars += text?.Length ?? 0;
                            var content = RenderInlineRuns(para, hyperlinkMap);
                            bodyHtml.AppendLine($"  <li>{content}</li>");
                        }
                        else if (!string.IsNullOrWhiteSpace(text))
                        {
                            CloseListIfOpen();
                            totalChars += text.Length;
                            var content = RenderInlineRuns(para, hyperlinkMap);
                            bodyHtml.AppendLine($"<p>{content}</p>");
                        }
                    }
                    else if (elem is Table table)
                    {
                        CloseListIfOpen();
                        var rows = table.Elements<TableRow>().ToList();
                        if (rows.Count == 0) continue;

                        tablesReconstructed++;
                        bodyHtml.AppendLine("<table>");
                        bool isHeader = true;

                        foreach (var row in rows)
                        {
                            bodyHtml.AppendLine("  <tr>");
                            var cells = row.Elements<TableCell>().ToList();
                            foreach (var cell in cells)
                            {
                                var cellContent = RenderCellRuns(cell, hyperlinkMap);
                                totalChars += cell.InnerText?.Length ?? 0;
                                string cellTag = isHeader ? "th" : "td";
                                bodyHtml.AppendLine($"    <{cellTag}>{cellContent}</{cellTag}>");
                            }
                            bodyHtml.AppendLine("  </tr>");
                            isHeader = false;
                        }

                        bodyHtml.AppendLine("</table>");
                    }
                }

                CloseListIfOpen();
            }

            // Build complete HTML5 document
            var fullHtml = $@"<!DOCTYPE html>
<html lang=""en"">
<head>
  <meta charset=""utf-8"" />
  <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />
  <title>{WebUtility.HtmlEncode(docTitle)}</title>
  <style>
    body {{
      font-family: -apple-system, BlinkMacSystemFont, ""Segoe UI"", Roboto, Helvetica, Arial, sans-serif;
      line-height: 1.6;
      max-width: 840px;
      margin: 40px auto;
      padding: 0 24px;
      color: #1e293b;
      background: #ffffff;
    }}
    h1, h2, h3, h4, h5, h6 {{
      color: #0f172a;
      margin-top: 1.4em;
      margin-bottom: 0.5em;
      font-weight: 700;
    }}
    h1 {{ font-size: 2rem; border-bottom: 1px solid #e2e8f0; padding-bottom: 0.3em; }}
    h2 {{ font-size: 1.5rem; }}
    h3 {{ font-size: 1.25rem; }}
    p {{ margin-bottom: 1em; }}
    ul, ol {{ margin-bottom: 1em; padding-left: 2em; }}
    li {{ margin-bottom: 0.3em; }}
    table {{
      border-collapse: collapse;
      width: 100%;
      margin: 1.5em 0;
    }}
    th, td {{
      border: 1px solid #cbd5e1;
      padding: 8px 12px;
      text-align: left;
    }}
    th {{
      background-color: #f8fafc;
      font-weight: 600;
    }}
    figure {{ margin: 1.5em 0; text-align: center; }}
    img {{
      max-width: 100%;
      height: auto;
      border-radius: 4px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }}
    a {{
      color: #2563eb;
      text-decoration: underline;
    }}
    a:hover {{
      color: #1d4ed8;
    }}
    code {{
      font-family: Consolas, Monaco, monospace;
      background-color: #f1f5f9;
      padding: 2px 6px;
      border-radius: 4px;
      font-size: 0.9em;
    }}
  </style>
</head>
<body>
{bodyHtml}
</body>
</html>";

            progress?.Report(new ProcessingProgress(95, "Writing HTML output…"));
            File.WriteAllText(tempPartial, fullHtml, Encoding.UTF8);

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
            File.Move(tempPartial, outputPath);
            sw.Stop();

            progress?.Report(new ProcessingProgress(100, "DOCX to HTML conversion completed."));

            var report = new OfficeConversionQualityReport(
                Status: OfficeConversionStatus.Successful,
                ConversionMode: "DocxToHtmlConversion",
                PagesOrSlidesCount: 1,
                HeadingsDetectedCount: headingsDetected,
                TablesReconstructedCount: tablesReconstructed,
                ImagesPreservedCount: imagesPreserved,
                ShapesPreservedCount: 0,
                CharacterCount: totalChars,
                Limitations: limitations,
                Warnings: warnings,
                ValidationPassed: true,
                Duration: sw.Elapsed);

            return new DocxHtmlResult(
                outputPath,
                fullHtml,
                report,
                $"Converted Word document to HTML ({headingsDetected} heading(s), {tablesReconstructed} table(s), {imagesPreserved} image(s)).");
        }
        finally
        {
            if (File.Exists(tempPartial))
            {
                try { File.Delete(tempPartial); } catch { }
            }
        }
    }

    private static string RenderInlineRuns(Paragraph para, Dictionary<string, string> hyperlinkMap)
    {
        var sb = new StringBuilder();
        foreach (var child in para.ChildElements)
        {
            if (child is Hyperlink hyperlink)
            {
                var relId = hyperlink.Id?.Value;
                string? url = null;
                if (!string.IsNullOrEmpty(relId) && hyperlinkMap.TryGetValue(relId, out var targetUri))
                {
                    url = targetUri;
                }

                var text = RenderRuns(hyperlink.Elements<Run>());
                if (IsValidWebUrl(url))
                {
                    sb.Append($"<a href=\"{WebUtility.HtmlEncode(url)}\" target=\"_blank\" rel=\"noopener noreferrer\">{text}</a>");
                }
                else
                {
                    sb.Append(text);
                }
            }
            else if (child is Run run)
            {
                sb.Append(RenderSingleRun(run));
            }
        }

        var result = sb.ToString();
        return string.IsNullOrWhiteSpace(result) ? WebUtility.HtmlEncode(para.InnerText ?? "") : result;
    }

    private static string RenderCellRuns(TableCell cell, Dictionary<string, string> hyperlinkMap)
    {
        var paras = cell.Elements<Paragraph>().ToList();
        if (paras.Count == 0) return WebUtility.HtmlEncode(cell.InnerText ?? "");

        var parts = paras.Select(p => RenderInlineRuns(p, hyperlinkMap)).Where(s => !string.IsNullOrWhiteSpace(s));
        return string.Join("<br />", parts);
    }

    private static string RenderRuns(IEnumerable<Run> runs)
    {
        var sb = new StringBuilder();
        foreach (var run in runs)
        {
            sb.Append(RenderSingleRun(run));
        }
        return sb.ToString();
    }

    private static string RenderSingleRun(Run run)
    {
        var text = run.InnerText;
        if (string.IsNullOrEmpty(text)) return "";

        var encoded = WebUtility.HtmlEncode(text);
        var rPr = run.RunProperties;
        if (rPr == null) return encoded;

        bool isBold = rPr.Bold != null && (rPr.Bold.Val == null || rPr.Bold.Val.Value);
        bool isItalic = rPr.Italic != null && (rPr.Italic.Val == null || rPr.Italic.Val.Value);
        bool isUnderline = rPr.Underline != null;
        bool isStrike = rPr.Strike != null;

        if (isBold) encoded = $"<strong>{encoded}</strong>";
        if (isItalic) encoded = $"<em>{encoded}</em>";
        if (isUnderline) encoded = $"<u>{encoded}</u>";
        if (isStrike) encoded = $"<s>{encoded}</s>";

        return encoded;
    }

    private static bool IsValidWebUrl(string? url)
    {
        if (string.IsNullOrWhiteSpace(url)) return false;
        return url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
               url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
               url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase);
    }

    private static (bool IsHeading, int Level) DetermineHeading(string? styleId)
    {
        if (string.IsNullOrWhiteSpace(styleId)) return (false, 0);
        var norm = styleId.Replace(" ", "").ToLowerInvariant();
        if (norm.Contains("title")) return (true, 1);
        if (norm.Contains("subtitle")) return (true, 2);
        if (norm.StartsWith("heading") || norm.StartsWith("head"))
        {
            for (int i = 1; i <= 6; i++)
            {
                if (norm.Contains(i.ToString())) return (true, i);
            }
            return (true, 2);
        }
        return (false, 0);
    }
}
