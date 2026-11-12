using System.Diagnostics;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using SmartShift.Core.IO;
using SmartShift.Processing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace SmartShift.Processing.Documents;

public static class PptxTextExtractor
{
    private const double EmuPerPoint = 12700.0;

    public static OfficeTextResult Extract(
        string sourcePptxPath,
        string slideRange = "all",
        string? outputPath = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(sourcePptxPath) || !File.Exists(sourcePptxPath))
        {
            throw new FileNotFoundException("Source PowerPoint presentation does not exist.", sourcePptxPath);
        }

        var sw = Stopwatch.StartNew();
        progress?.Report(new ProcessingProgress(15, "Opening PowerPoint presentation package…"));

        outputPath ??= Path.ChangeExtension(sourcePptxPath, ".txt");
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

        int slidesProcessed = 0;
        int headingsDetected = 0;
        int tablesReconstructed = 0;
        var sb = new StringBuilder();
        var limitations = new List<string>();
        var warnings = new List<string>();

        try
        {
            using (var presentationDoc = PresentationDocument.Open(sourcePptxPath, false))
            {
                var presentationPart = presentationDoc.PresentationPart;
                if (presentationPart?.Presentation == null)
                {
                    throw new InvalidOperationException("Invalid presentation: missing PresentationPart or Presentation root.");
                }

                var slideIds = presentationPart.Presentation.SlideIdList?.Elements<P.SlideId>().ToList() ?? [];
                if (slideIds.Count == 0)
                {
                    throw new InvalidOperationException("The PowerPoint presentation contains no slides.");
                }

                var selectedIndices = PageRangeParser.Parse(slideRange, slideIds.Count);
                if (selectedIndices.Count == 0)
                {
                    throw new InvalidOperationException($"No slides matched the specified range '{slideRange}'.");
                }

                for (int i = 0; i < selectedIndices.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int slideIdx = selectedIndices[i] - 1;
                    int pct = 20 + (int)(70.0 * i / selectedIndices.Count);
                    progress?.Report(new ProcessingProgress(pct, $"Extracting text from slide {i + 1} of {selectedIndices.Count}…"));

                    var slideId = slideIds[slideIdx];
                    var relId = slideId.RelationshipId?.Value;
                    if (string.IsNullOrEmpty(relId) || !presentationPart.TryGetPartById(relId, out var part) || part is not SlidePart slidePart)
                    {
                        warnings.Add($"Slide {selectedIndices[i]} could not be resolved from presentation package.");
                        continue;
                    }

                    slidesProcessed++;
                    var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
                    if (shapeTree == null) continue;

                    // 1. Identify title
                    string? slideTitle = null;
                    var titleShape = shapeTree.Elements<P.Shape>().FirstOrDefault(s =>
                    {
                        var phType = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value;
                        return phType == P.PlaceholderValues.Title || phType == P.PlaceholderValues.CenteredTitle;
                    });

                    if (titleShape?.TextBody != null)
                    {
                        slideTitle = titleShape.TextBody.InnerText?.Trim();
                    }

                    if (!string.IsNullOrWhiteSpace(slideTitle))
                    {
                        headingsDetected++;
                        sb.AppendLine($"--- Slide {selectedIndices[i]}: {slideTitle} ---");
                    }
                    else
                    {
                        sb.AppendLine($"--- Slide {selectedIndices[i]} ---");
                    }
                    sb.AppendLine();

                    // 2. Extract shape elements ordered vertically
                    var shapes = shapeTree.Elements<P.Shape>()
                        .Where(s => s != titleShape)
                        .Select(s => new
                        {
                            Shape = s,
                            Y = (s.ShapeProperties?.Transform2D?.Offset?.Y?.Value ?? 0L) / EmuPerPoint,
                            X = (s.ShapeProperties?.Transform2D?.Offset?.X?.Value ?? 0L) / EmuPerPoint
                        })
                        .OrderBy(item => item.Y)
                        .ThenBy(item => item.X)
                        .ToList();

                    foreach (var item in shapes)
                    {
                        var textBody = item.Shape.TextBody;
                        if (textBody == null) continue;

                        foreach (var para in textBody.Elements<A.Paragraph>())
                        {
                            var text = para.InnerText?.Trim();
                            if (string.IsNullOrWhiteSpace(text)) continue;

                            bool isBullet = para.ParagraphProperties?.Descendants<A.CharacterBullet>().Any() == true ||
                                            para.ParagraphProperties?.Descendants<A.AutoNumberedBullet>().Any() == true;

                            if (isBullet)
                            {
                                sb.AppendLine($"• {text}");
                            }
                            else
                            {
                                sb.AppendLine(text);
                            }
                        }
                        sb.AppendLine();
                    }

                    // 3. Extract tables
                    var graphicFrames = shapeTree.Elements<P.GraphicFrame>().ToList();
                    foreach (var frame in graphicFrames)
                    {
                        var table = frame.Descendants<A.Table>().FirstOrDefault();
                        if (table == null) continue;

                        var rows = table.Elements<A.TableRow>().ToList();
                        if (rows.Count == 0) continue;

                        tablesReconstructed++;
                        bool isHeader = true;
                        foreach (var row in rows)
                        {
                            var cells = row.Elements<A.TableCell>().ToList();
                            var cellTexts = cells.Select(c => (c.TextBody?.InnerText?.Trim() ?? "").Replace("\r", " ").Replace("\n", " ")).ToList();
                            sb.AppendLine($"| {string.Join(" | ", cellTexts)} |");

                            if (isHeader)
                            {
                                sb.AppendLine($"| {string.Join(" | ", cellTexts.Select(_ => "---"))} |");
                                isHeader = false;
                            }
                        }
                        sb.AppendLine();
                    }

                    sb.AppendLine();
                }
            }

            var fullText = sb.ToString().Trim();
            progress?.Report(new ProcessingProgress(95, "Writing text output…"));

            File.WriteAllText(tempPartial, fullText, Encoding.UTF8);

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
            File.Move(tempPartial, outputPath);
            sw.Stop();

            progress?.Report(new ProcessingProgress(100, "PPTX text extraction completed."));

            var report = new OfficeConversionQualityReport(
                Status: OfficeConversionStatus.Successful,
                ConversionMode: "PptxTextExtraction",
                PagesOrSlidesCount: slidesProcessed,
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
                $"Extracted {fullText.Length} character(s) from {slidesProcessed} slide(s) ({headingsDetected} title(s), {tablesReconstructed} table(s)).");
        }
        finally
        {
            if (File.Exists(tempPartial))
            {
                try { File.Delete(tempPartial); } catch { }
            }
        }
    }
}
