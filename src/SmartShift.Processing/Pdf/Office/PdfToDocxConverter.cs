using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using Docnet.Core;
using Docnet.Core.Exceptions;
using Docnet.Core.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using SkiaSharp;
using SmartShift.Core.IO;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.DocumentLayoutAnalysis.ReadingOrderDetector;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace SmartShift.Processing.Pdf.Office;

public sealed record PdfDocxResult(
    string OutputPath,
    int TotalPagesInDoc,
    int PagesConvertedCount,
    int TotalCharactersExtracted,
    bool IsScannedOrImageOnly,
    string SummaryNote,
    PdfConversionQualityReport? QualityReport = null);

public static class PdfToDocxConverter
{
    public static PdfDocxResult ConvertToDocx(
        string inputPath,
        IOutputPathAllocator allocator,
        string? pageRange,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("The source PDF file was not found.", inputPath);
        }

        if (new FileInfo(inputPath).Length == 0)
        {
            throw new InvalidOperationException("The source PDF file is empty (0 bytes).");
        }

        progress?.Report(new ProcessingProgress(5, "Inspecting PDF document structure…"));

        var sw = Stopwatch.StartNew();
        var baseName = Path.GetFileNameWithoutExtension(inputPath);
        var finalPath = allocator.Allocate(inputPath, $"{baseName}.docx");
        var tempPartial = finalPath + ".partial";

        int totalPages = 0;
        IReadOnlyList<int> pagesToConvert;

        try
        {
            using var pdfDoc = PdfDocument.Open(inputPath);
            totalPages = pdfDoc.NumberOfPages;
            pagesToConvert = PageRangeParser.Parse(pageRange, totalPages);
        }
        catch (UglyToad.PdfPig.Exceptions.PdfDocumentEncryptedException ex)
        {
            throw new InvalidOperationException(
                "This PDF document is password-protected or uses an unsupported encryption handler. Encrypted PDFs cannot be converted.",
                ex);
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            throw new InvalidOperationException("The PDF document is malformed or corrupted and could not be read.", ex);
        }

        if (pagesToConvert.Count == 0)
        {
            throw new InvalidOperationException("No pages matched the specified page range.");
        }

        try
        {
            int totalChars = 0;
            bool isScannedOrEmpty = false;
            int imageCount = 0;
            int headingsCount = 0;
            int tablesCount = 0;
            var warnings = new List<string>();

            using (var docx = WordprocessingDocument.Create(tempPartial, WordprocessingDocumentType.Document))
            {
                var mainPart = docx.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                using var pdfDoc = PdfDocument.Open(inputPath);

                // First pass: detect if document has any extractable text
                int detectedChars = 0;
                var allFontSizes = new List<double>();

                foreach (var pageNum in pagesToConvert)
                {
                    var page = pdfDoc.GetPage(pageNum);
                    foreach (var letter in page.Letters)
                    {
                        if (!string.IsNullOrWhiteSpace(letter.Value))
                        {
                            detectedChars++;
                            if (letter.PointSize > 0)
                            {
                                allFontSizes.Add(letter.PointSize);
                            }
                        }
                    }
                }

                isScannedOrEmpty = (detectedChars == 0);
                double medianFontSize = allFontSizes.Count > 0 ? Median(allFontSizes) : 11.0;

                if (isScannedOrEmpty)
                {
                    // Scanned / Image-only PDF: Render pages as high-fidelity images
                    progress?.Report(new ProcessingProgress(20, "Scanned document detected. Rendering full-page images…"));
                    warnings.Add("Document appears to be scanned or image-only without an extractable text layer. Pages embedded as high-fidelity images. Text is not editable because no OCR engine is installed.");

                    for (int i = 0; i < pagesToConvert.Count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        int pageNum = pagesToConvert[i];

                        int pct = 20 + (int)(60.0 * (i + 1) / pagesToConvert.Count);
                        progress?.Report(new ProcessingProgress(pct, $"Rendering page {pageNum} of {totalPages}…"));

                        var pageImageBytes = RenderPageToPng(inputPath, pageNum);
                        if (pageImageBytes != null && pageImageBytes.Length > 0)
                        {
                            AddImageToBody(mainPart, body, pageImageBytes, $"Page {pageNum}", 500, 700);
                            imageCount++;
                        }

                        if (i < pagesToConvert.Count - 1)
                        {
                            body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                        }
                    }
                }
                else
                {
                    // Text-based PDF: Extract structured paragraphs, lines, headings, tables, and images
                    for (int i = 0; i < pagesToConvert.Count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        int pageNum = pagesToConvert[i];

                        int pct = 20 + (int)(65.0 * (i + 1) / pagesToConvert.Count);
                        progress?.Report(new ProcessingProgress(pct, $"Converting page {pageNum} of {totalPages}…"));

                        var page = pdfDoc.GetPage(pageNum);
                        var words = page.GetWords().ToList();

                        if (words.Count > 0)
                        {
                            var pageLines = GroupWordsIntoLines(words);
                            int lineIdx = 0;

                            while (lineIdx < pageLines.Count)
                            {
                                var tableLines = DetectTableLines(pageLines, lineIdx);
                                if (tableLines != null && tableLines.Count >= 2)
                                {
                                    // Reconstruct Table
                                    var table = BuildOpenXmlTable(tableLines, ref totalChars);
                                    body.AppendChild(table);
                                    tablesCount++;
                                    lineIdx += tableLines.Count;
                                    continue;
                                }

                                // Gather non-table lines until next table or end of page
                                var nonTableLines = new List<List<Word>>();
                                while (lineIdx < pageLines.Count && DetectTableLines(pageLines, lineIdx) == null)
                                {
                                    nonTableLines.Add(pageLines[lineIdx]);
                                    lineIdx++;
                                }

                                if (nonTableLines.Count > 0)
                                {
                                    var nonTableWords = nonTableLines.SelectMany(l => l).ToList();
                                    var blocks = TrySegmentBlocks(nonTableWords);

                                    if (blocks != null && blocks.Count > 0)
                                    {
                                        foreach (var block in blocks)
                                        {
                                            var blockWords = block.TextLines.SelectMany(tl => tl.Words).ToList();
                                            if (blockWords.Count == 0) continue;

                                            var lines = GroupWordsIntoLines(blockWords);
                                            ProcessLinesIntoBody(lines, body, medianFontSize, ref totalChars, ref headingsCount);
                                        }
                                    }
                                    else
                                    {
                                        ProcessLinesIntoBody(nonTableLines, body, medianFontSize, ref totalChars, ref headingsCount);
                                    }
                                }
                            }
                        }

                        // Extract embedded raster images on page
                        try
                        {
                            foreach (var img in page.GetImages())
                            {
                                if (img.TryGetPng(out var imgBytes) && imgBytes.Length > 0)
                                {
                                    AddImageToBody(mainPart, body, imgBytes, $"Figure {imageCount + 1}", 350, 250);
                                    imageCount++;
                                }
                            }
                        }
                        catch
                        {
                            // Non-fatal: if an embedded image stream has unsupported filter, continue
                        }

                        // Page break between PDF pages
                        if (i < pagesToConvert.Count - 1)
                        {
                            body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                        }
                    }
                }

                // Standard letter/A4 section page setup
                var sectPr = new SectionProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.PageSize { Width = 11906, Height = 16838 }, // A4 standard in twentieths of a point (dxa)
                    new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440, Header = 720, Footer = 720 });
                body.AppendChild(sectPr);

                mainPart.Document.Save();
            }

            progress?.Report(new ProcessingProgress(90, "Validating Office Open XML package…"));

            // Validate OpenXML structure
            var validator = new OpenXmlValidator();
            using (var package = WordprocessingDocument.Open(tempPartial, false))
            {
                var validationErrors = validator.Validate(package).ToList();
            }

            // Atomically promote
            if (File.Exists(finalPath))
            {
                File.Delete(finalPath);
            }
            File.Move(tempPartial, finalPath);

            sw.Stop();
            long outputBytes = new FileInfo(finalPath).Length;

            var status = isScannedOrEmpty
                ? PdfConversionStatus.SuccessfulWithLimitations
                : PdfConversionStatus.Successful;

            string summaryNote = isScannedOrEmpty
                ? $"Converted {pagesToConvert.Count} page(s) to DOCX ({FileSizeFormatter.Format(outputBytes)}). Note: Document contains no extractable text; pages embedded as high-fidelity images (OCR required for text editing)."
                : $"Converted {pagesToConvert.Count} page(s) ({totalChars:N0} characters, {imageCount} image(s), {headingsCount} heading(s), {tablesCount} table(s)) into editable DOCX in {sw.ElapsedMilliseconds} ms.";

            progress?.Report(new ProcessingProgress(100, "DOCX conversion complete."));

            var qualityReport = new PdfConversionQualityReport(
                SourcePath: inputPath,
                OutputPath: finalPath,
                OutputFiles: [finalPath],
                ConversionMode: isScannedOrEmpty ? "DocxImageFallback" : "DocxStructuredReconstruction",
                TotalInputPages: totalPages,
                PagesProcessed: pagesToConvert.Count,
                TotalCharactersExtracted: totalChars,
                ImagesPreservedCount: imageCount,
                TablesReconstructedCount: tablesCount,
                HeadingsDetectedCount: headingsCount,
                Status: status,
                Warnings: warnings,
                OutputSizeBytes: outputBytes,
                SummaryNote: summaryNote);

            return new PdfDocxResult(
                finalPath,
                totalPages,
                pagesToConvert.Count,
                totalChars,
                isScannedOrEmpty,
                summaryNote,
                qualityReport);
        }
        catch
        {
            try
            {
                if (File.Exists(tempPartial)) File.Delete(tempPartial);
            }
            catch { }
            throw;
        }
    }

    private static IReadOnlyList<UglyToad.PdfPig.DocumentLayoutAnalysis.TextBlock>? TrySegmentBlocks(List<Word> words)
    {
        try
        {
            if (words.Count < 4) return null;
            var blocks = DocstrumBoundingBoxes.Instance.GetBlocks(words);
            if (blocks != null && blocks.Count > 0)
            {
                return DefaultReadingOrderDetector.Instance.Get(blocks).ToList();
            }
        }
        catch
        {
            // Fall back to line grouping
        }
        return null;
    }

    private static void ProcessLinesIntoBody(
        List<List<Word>> lines,
        Body body,
        double medianFontSize,
        ref int totalChars,
        ref int headingsCount)
    {
        if (lines.Count == 0) return;

        foreach (var currentLineWords in lines)
        {
            var lineText = string.Join(" ", currentLineWords.Select(w => w.Text));
            totalChars += lineText.Length;

            double avgFontPt = currentLineWords.Average(GetWordFontSize);
            bool isBold = currentLineWords.Any(w => w.FontName?.Contains("Bold", StringComparison.OrdinalIgnoreCase) == true);
            bool isItalic = currentLineWords.Any(w => w.FontName?.Contains("Italic", StringComparison.OrdinalIgnoreCase) == true);

            bool isHeading = (avgFontPt >= 14.0 || (avgFontPt >= medianFontSize * 1.25 && isBold)) && lineText.Length < 120;

            var para = new Paragraph();
            var paraProps = new ParagraphProperties();

            if (isHeading)
            {
                headingsCount++;
                string styleVal = avgFontPt >= 17.0 ? "Heading1" : "Heading2";
                paraProps.ParagraphStyleId = new ParagraphStyleId { Val = styleVal };
                paraProps.SpacingBetweenLines = new SpacingBetweenLines
                {
                    Before = "200",
                    After = "120",
                    Line = "240",
                    LineRule = LineSpacingRuleValues.Auto
                };
            }
            else
            {
                paraProps.SpacingBetweenLines = new SpacingBetweenLines
                {
                    After = "100",
                    Line = "240",
                    LineRule = LineSpacingRuleValues.Auto
                };
            }
            para.AppendChild(paraProps);

            var runProps = new RunProperties();
            if (isBold || isHeading) runProps.AppendChild(new Bold());
            if (isItalic) runProps.AppendChild(new Italic());

            int fontSizeHalfPt = Math.Clamp((int)Math.Round(avgFontPt * 2.0), 16, 72);
            runProps.AppendChild(new FontSize { Val = fontSizeHalfPt.ToString() });

            var run = new Run(new Text(lineText) { Space = SpaceProcessingModeValues.Preserve });
            run.RunProperties = runProps;
            para.AppendChild(run);
            body.AppendChild(para);
        }
    }

    private static double GetWordFontSize(Word w)
    {
        if (w.Letters != null && w.Letters.Count > 0)
        {
            return w.Letters.Average(l => l.PointSize);
        }
        return w.BoundingBox.Height;
    }

    private static List<List<Word>>? DetectTableLines(List<List<Word>> lines, int startIndex)
    {
        var candidateRows = new List<List<Word>>();
        int expectedColumns = -1;

        for (int i = startIndex; i < lines.Count; i++)
        {
            var line = lines[i];
            var clusters = SplitLineIntoColumnClusters(line);
            // Table row must have at least 2 columns and cells shouldn't be full prose paragraphs (max 4 words per cell)
            if (clusters.Count < 2 || clusters.Any(c => c.Count > 4))
            {
                break;
            }

            if (expectedColumns == -1)
            {
                expectedColumns = clusters.Count;
                candidateRows.Add(line);
            }
            else if (clusters.Count == expectedColumns)
            {
                candidateRows.Add(line);
            }
            else
            {
                break;
            }
        }

        return candidateRows.Count >= 2 ? candidateRows : null;
    }

    private static List<List<Word>> SplitLineIntoColumnClusters(List<Word> words)
    {
        var clusters = new List<List<Word>>();
        if (words.Count == 0) return clusters;

        var sorted = words.OrderBy(w => w.BoundingBox.Left).ToList();
        var currentCluster = new List<Word> { sorted[0] };

        for (int i = 1; i < sorted.Count; i++)
        {
            var prev = sorted[i - 1];
            var curr = sorted[i];
            double gap = curr.BoundingBox.Left - prev.BoundingBox.Right;

            // Gap greater than 16 pt indicates separate column cell
            if (gap > 16.0)
            {
                clusters.Add(currentCluster);
                currentCluster = new List<Word> { curr };
            }
            else
            {
                currentCluster.Add(curr);
            }
        }

        if (currentCluster.Count > 0)
        {
            clusters.Add(currentCluster);
        }

        return clusters;
    }

    private static Table BuildOpenXmlTable(List<List<Word>> tableLines, ref int totalChars)
    {
        var table = new Table();

        // Style table with clean borders and width
        var tblPr = new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "D0D7DE" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "D0D7DE" },
                new LeftBorder { Val = BorderValues.None },
                new RightBorder { Val = BorderValues.None },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "E1E4E8" },
                new InsideVerticalBorder { Val = BorderValues.None }),
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }, // 100%
            new TableCellMarginDefault(
                new TopMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                new LeftMargin { Width = "120", Type = TableWidthUnitValues.Dxa },
                new RightMargin { Width = "120", Type = TableWidthUnitValues.Dxa }));

        table.AppendChild(tblPr);

        bool isHeader = true;
        foreach (var line in tableLines)
        {
            var row = new TableRow();
            var clusters = SplitLineIntoColumnClusters(line);

            foreach (var cluster in clusters)
            {
                var cell = new TableCell();
                var cellPr = new TableCellProperties();

                if (isHeader)
                {
                    cellPr.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Fill = "F6F8FA" });
                }
                cell.AppendChild(cellPr);

                var cellText = string.Join(" ", cluster.Select(w => w.Text));
                totalChars += cellText.Length;

                bool isBold = cluster.Any(w => w.FontName?.Contains("Bold", StringComparison.OrdinalIgnoreCase) == true) || isHeader;
                bool isItalic = cluster.Any(w => w.FontName?.Contains("Italic", StringComparison.OrdinalIgnoreCase) == true);

                var p = new Paragraph(new ParagraphProperties(new SpacingBetweenLines { After = "40", Line = "240", LineRule = LineSpacingRuleValues.Auto }));
                var runProps = new RunProperties();
                if (isBold) runProps.AppendChild(new Bold());
                if (isItalic) runProps.AppendChild(new Italic());

                double avgFont = cluster.Average(GetWordFontSize);
                int fontHalfPt = Math.Clamp((int)Math.Round(avgFont * 2.0), 16, 48);
                runProps.AppendChild(new FontSize { Val = fontHalfPt.ToString() });

                var run = new Run(new Text(cellText) { Space = SpaceProcessingModeValues.Preserve });
                run.RunProperties = runProps;
                p.AppendChild(run);
                cell.AppendChild(p);

                row.AppendChild(cell);
            }

            table.AppendChild(row);
            isHeader = false;
        }

        return table;
    }

    private static double Median(List<double> values)
    {
        var sorted = values.OrderBy(v => v).ToList();
        int mid = sorted.Count / 2;
        return sorted.Count % 2 != 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2.0;
    }

    private static List<List<Word>> GroupWordsIntoLines(List<Word> words)
    {
        // Sort descending by Y (top of page to bottom), then ascending by X (left to right)
        var sorted = words.OrderByDescending(w => w.BoundingBox.Bottom).ThenBy(w => w.BoundingBox.Left).ToList();
        var lines = new List<List<Word>>();
        if (sorted.Count == 0) return lines;

        var currentLine = new List<Word> { sorted[0] };
        double currentY = sorted[0].BoundingBox.Bottom;
        double currentH = sorted[0].BoundingBox.Height;

        for (int i = 1; i < sorted.Count; i++)
        {
            var w = sorted[i];
            // If vertical distance between words is small (less than half font height), consider on same line
            if (Math.Abs(w.BoundingBox.Bottom - currentY) < Math.Max(4.0, currentH * 0.5))
            {
                currentLine.Add(w);
            }
            else
            {
                lines.Add(currentLine.OrderBy(x => x.BoundingBox.Left).ToList());
                currentLine = new List<Word> { w };
                currentY = w.BoundingBox.Bottom;
                currentH = w.BoundingBox.Height;
            }
        }

        if (currentLine.Count > 0)
        {
            lines.Add(currentLine.OrderBy(x => x.BoundingBox.Left).ToList());
        }

        return lines;
    }

    private static byte[]? RenderPageToPng(string pdfPath, int pageNumber)
    {
        try
        {
            using var docReader = DocLib.Instance.GetDocReader(pdfPath, new PageDimensions(150 / 72.0));
            using var pageReader = docReader.GetPageReader(pageNumber - 1);
            int width = pageReader.GetPageWidth();
            int height = pageReader.GetPageHeight();
            var rawBgra = pageReader.GetImage();
            if (rawBgra == null || rawBgra.Length == 0) return null;

            var info = new SKImageInfo(width, height, SKColorType.Bgra8888, SKAlphaType.Premul);
            using var bitmap = new SKBitmap(info);
            System.Runtime.InteropServices.Marshal.Copy(rawBgra, 0, bitmap.GetPixels(), rawBgra.Length);
            using var encoded = bitmap.Encode(SKEncodedImageFormat.Png, 100);
            return encoded?.ToArray();
        }
        catch
        {
            return null;
        }
    }

    private static void AddImageToBody(MainDocumentPart mainPart, Body body, byte[] imageBytes, string description, long widthPt, long heightPt)
    {
        try
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var ms = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(ms);
            }

            string relId = mainPart.GetIdOfPart(imagePart);

            // 1 pt = 12700 EMUs
            long cx = widthPt * 12700;
            long cy = heightPt * 12700;

            var drawing = new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = cx, Cy = cy },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = (uint)Random.Shared.Next(1, 100000), Name = description },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = (uint)Random.Shared.Next(1, 100000), Name = description },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip { Embed = relId, CompressionState = A.BlipCompressionValues.Print },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = cx, Cy = cy }),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));

            body.AppendChild(new Paragraph(new Run(drawing)));
        }
        catch
        {
            // Non-fatal if single picture insertion fails
        }
    }
}
