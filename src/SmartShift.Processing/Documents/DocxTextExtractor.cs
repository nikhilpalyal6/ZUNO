using System.Diagnostics;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SmartShift.Processing;

namespace SmartShift.Processing.Documents;

public static class DocxTextExtractor
{
    public static OfficeTextResult Extract(
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
        var sb = new StringBuilder();
        var limitations = new List<string>();
        var warnings = new List<string>();

        try
        {
            using (var wordDoc = WordprocessingDocument.Open(sourceDocxPath, false))
            {
                var body = wordDoc.MainDocumentPart?.Document?.Body;
                if (body == null)
                {
                    throw new InvalidOperationException("Word document does not contain a valid main body.");
                }

                var elements = body.ChildElements.ToList();
                int totalElements = elements.Count;

                for (int i = 0; i < totalElements; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (i % 20 == 0)
                    {
                        int pct = 20 + (int)(70.0 * i / Math.Max(1, totalElements));
                        progress?.Report(new ProcessingProgress(pct, $"Extracting document content ({i + 1}/{totalElements})…"));
                    }

                    var elem = elements[i];

                    if (elem is Paragraph para)
                    {
                        // Check for explicit page break in runs
                        bool hasPageBreak = para.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page) ||
                                            para.Descendants<LastRenderedPageBreak>().Any();

                        if (hasPageBreak)
                        {
                            if (sb.Length > 0 && !sb.ToString().EndsWith("\n\n", StringComparison.Ordinal))
                            {
                                sb.AppendLine();
                            }
                            sb.AppendLine("--- Page Break ---");
                            sb.AppendLine();
                        }

                        var text = para.InnerText?.Trim();
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            continue;
                        }

                        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                        var (isHeading, level) = DetermineHeading(styleId);
                        bool isBullet = para.ParagraphProperties?.NumberingProperties != null;

                        if (isHeading)
                        {
                            headingsDetected++;
                            string prefix = level == 1 ? "# " : (level == 2 ? "## " : "### ");
                            sb.AppendLine(prefix + text);
                            sb.AppendLine();
                        }
                        else if (isBullet)
                        {
                            sb.AppendLine($"• {text}");
                        }
                        else
                        {
                            sb.AppendLine(text);
                            sb.AppendLine();
                        }
                    }
                    else if (elem is Table table)
                    {
                        var rows = table.Elements<TableRow>().ToList();
                        if (rows.Count == 0) continue;

                        tablesReconstructed++;
                        if (sb.Length > 0 && !sb.ToString().EndsWith("\n\n", StringComparison.Ordinal))
                        {
                            sb.AppendLine();
                        }

                        bool isHeader = true;
                        foreach (var row in rows)
                        {
                            var cells = row.Elements<TableCell>().ToList();
                            var cellTexts = cells.Select(c => (c.InnerText?.Trim() ?? "").Replace("\r", " ").Replace("\n", " ")).ToList();
                            sb.AppendLine($"| {string.Join(" | ", cellTexts)} |");

                            if (isHeader)
                            {
                                sb.AppendLine($"| {string.Join(" | ", cellTexts.Select(_ => "---"))} |");
                                isHeader = false;
                            }
                        }

                        sb.AppendLine();
                    }
                }
            }

            var fullText = sb.ToString().Trim();
            progress?.Report(new ProcessingProgress(92, "Writing plain text output…"));

            File.WriteAllText(tempPartial, fullText, Encoding.UTF8);

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
            File.Move(tempPartial, outputPath);
            sw.Stop();

            progress?.Report(new ProcessingProgress(100, "DOCX text extraction completed."));

            var report = new OfficeConversionQualityReport(
                Status: OfficeConversionStatus.Successful,
                ConversionMode: "DocxTextExtraction",
                PagesOrSlidesCount: 1,
                HeadingsDetectedCount: headingsDetected,
                TablesReconstructedCount: tablesReconstructed,
                ImagesPreservedCount: 0,
                ShapesPreservedCount: 0,
                CharacterCount: fullText.Length,
                Limitations: limitations,
                Warnings: warnings,
                ValidationPassed: true,
                Duration: sw.Elapsed);

            return new OfficeTextResult(
                outputPath,
                fullText,
                report,
                $"Extracted {fullText.Length} character(s) from Word document ({headingsDetected} heading(s), {tablesReconstructed} table(s)).");
        }
        finally
        {
            if (File.Exists(tempPartial))
            {
                try { File.Delete(tempPartial); } catch { }
            }
        }
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
