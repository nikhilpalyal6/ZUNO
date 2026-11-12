using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using SkiaSharp;
using SmartShift.Core.IO;
using SmartShift.Processing;
using A = DocumentFormat.OpenXml.Drawing;

namespace SmartShift.Processing.Documents;

public sealed record DocxPdfConversionOptions(
    bool PreferLocalWordEngine = false,
    string PageSize = "A4",
    double MarginPoints = 45.0,
    bool IncludePageNumbers = true);

public static class DocxToPdfConverter
{
    private const double PointsPerInch = 72.0;

    public static string Convert(
        string sourceDocxPath,
        string outputPath,
        DocxPdfConversionOptions? options = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        return ConvertWithReport(sourceDocxPath, outputPath, options, progress, cancellationToken).OutputPath;
    }

    public static OfficeConversionResult ConvertWithReport(
        string sourceDocxPath,
        string outputPath,
        DocxPdfConversionOptions? options = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(sourceDocxPath) || !File.Exists(sourceDocxPath))
        {
            throw new FileNotFoundException("Source Word document does not exist.", sourceDocxPath);
        }

        var sourceInfo = new FileInfo(sourceDocxPath);
        if (sourceInfo.Length == 0)
        {
            throw new InvalidOperationException("Source Word document is empty (0 bytes).");
        }

        var sw = Stopwatch.StartNew();
        options ??= new DocxPdfConversionOptions();

        var dir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        var tempPartial = outputPath + ".partial";
        if (File.Exists(tempPartial))
        {
            try { File.Delete(tempPartial); } catch { }
        }

        int pagesPreserved = 0;
        int headingsDetected = 0;
        int tablesReconstructed = 0;
        int imagesPreserved = 0;
        long totalChars = 0;
        var limitations = new List<string>();
        var warnings = new List<string>();

        // 1. If PreferLocalWordEngine is requested, try Word COM if available
        if (options.PreferLocalWordEngine)
        {
            try
            {
                if (TryConvertWithWordCom(sourceDocxPath, tempPartial, progress, cancellationToken))
                {
                    FinalizeAndValidate(tempPartial, outputPath, progress);
                    sw.Stop();
                    var comReport = new OfficeConversionQualityReport(
                        Status: OfficeConversionStatus.Successful,
                        ConversionMode: "NativeWordComEngine",
                        PagesOrSlidesCount: 1,
                        HeadingsDetectedCount: 0,
                        TablesReconstructedCount: 0,
                        ImagesPreservedCount: 0,
                        ShapesPreservedCount: 0,
                        CharacterCount: sourceInfo.Length,
                        Limitations: limitations,
                        Warnings: warnings,
                        ValidationPassed: true,
                        Duration: sw.Elapsed);

                    return new OfficeConversionResult(outputPath, comReport, "Converted via Microsoft Word native engine.");
                }
            }
            catch (Exception)
            {
                // Fall back cleanly to built-in OpenXML renderer
            }
        }

        // 2. Built-in OpenXML -> PDF renderer
        ConvertWithBuiltinEngine(sourceDocxPath, tempPartial, options, progress, cancellationToken,
            ref pagesPreserved, ref headingsDetected, ref tablesReconstructed, ref imagesPreserved, ref totalChars, limitations, warnings);

        FinalizeAndValidate(tempPartial, outputPath, progress);
        sw.Stop();

        var status = limitations.Count > 0 ? OfficeConversionStatus.SuccessfulWithLimitations : OfficeConversionStatus.Successful;
        var report = new OfficeConversionQualityReport(
            Status: status,
            ConversionMode: "DocxToPdfConversion",
            PagesOrSlidesCount: pagesPreserved,
            HeadingsDetectedCount: headingsDetected,
            TablesReconstructedCount: tablesReconstructed,
            ImagesPreservedCount: imagesPreserved,
            ShapesPreservedCount: 0,
            CharacterCount: totalChars,
            Limitations: limitations,
            Warnings: warnings,
            ValidationPassed: true,
            Duration: sw.Elapsed);

        return new OfficeConversionResult(
            outputPath,
            report,
            $"Converted Word document to PDF ({pagesPreserved} page(s), {headingsDetected} heading(s), {tablesReconstructed} table(s)).");
    }

    private static void FinalizeAndValidate(string tempPartial, string outputPath, IProgress<ProcessingProgress>? progress)
    {
        progress?.Report(new ProcessingProgress(95, "Validating output PDF integrity…"));
        if (!File.Exists(tempPartial) || new FileInfo(tempPartial).Length == 0)
        {
            throw new InvalidOperationException("Output PDF document could not be generated.");
        }

        using (var verifyStream = new FileStream(tempPartial, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (var verifyDoc = PdfReader.Open(verifyStream, PdfDocumentOpenMode.Import))
        {
            if (verifyDoc.PageCount <= 0)
            {
                throw new InvalidOperationException("Generated PDF contains no pages.");
            }
        }

        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }
        File.Move(tempPartial, outputPath);
        progress?.Report(new ProcessingProgress(100, "DOCX to PDF conversion completed."));
    }

    private static bool TryConvertWithWordCom(
        string sourceDocxPath,
        string targetPdfPath,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!OperatingSystem.IsWindows()) return false;
        Type? wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType is null) return false;

        progress?.Report(new ProcessingProgress(20, "Launching local Word engine for native conversion…"));
        dynamic? wordApp = null;
        dynamic? doc = null;
        try
        {
            wordApp = Activator.CreateInstance(wordType);
            if (wordApp is null) return false;

            wordApp.Visible = false;
            wordApp.DisplayAlerts = 0; // wdAlertsNone

            progress?.Report(new ProcessingProgress(45, "Opening document in Word engine…"));
            doc = wordApp.Documents.Open(
                FileName: Path.GetFullPath(sourceDocxPath),
                ReadOnly: true,
                ConfirmConversions: false,
                AddToRecentFiles: false,
                Visible: false);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ProcessingProgress(75, "Exporting native PDF from Word…"));

            // 17 = wdExportFormatPDF
            doc.ExportAsFixedFormat(
                OutputFileName: Path.GetFullPath(targetPdfPath),
                ExportFormat: 17,
                OpenAfterExport: false,
                OptimizeFor: 0); // wdExportOptimizeForPrint

            return File.Exists(targetPdfPath) && new FileInfo(targetPdfPath).Length > 0;
        }
        finally
        {
            try { doc?.Close(SaveChanges: 0); } catch { }
            try { wordApp?.Quit(SaveChanges: 0); } catch { }
        }
    }

    private static void ConvertWithBuiltinEngine(
        string sourceDocxPath,
        string tempPartial,
        DocxPdfConversionOptions options,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken,
        ref int pagesPreserved,
        ref int headingsDetected,
        ref int tablesReconstructed,
        ref int imagesPreserved,
        ref long totalChars,
        List<string> limitations,
        List<string> warnings)
    {
        cancellationToken.ThrowIfCancellationRequested();
        progress?.Report(new ProcessingProgress(15, "Reading Word document structure…"));

        using var wordDoc = WordprocessingDocument.Open(sourceDocxPath, false);
        var body = wordDoc.MainDocumentPart?.Document?.Body;
        if (body is null)
        {
            throw new InvalidOperationException("The Word document does not contain a valid document body.");
        }

        var mainPart = wordDoc.MainDocumentPart!;

        // Check unsupported elements
        if (body.Descendants().Any(e => e.LocalName == "oMath"))
        {
            limitations.Add("Mathematical equations (oMath) rendered as plain text or skipped.");
        }
        if (body.Descendants().Any(e => e.LocalName == "chart"))
        {
            limitations.Add("Embedded Office charts skipped or rendered as outline placeholders.");
        }

        // Section Properties inspection
        var sectPr = body.Elements<SectionProperties>().FirstOrDefault() ?? body.Descendants<SectionProperties>().FirstOrDefault();
        double pageWidth = 595.0; // A4 pt
        double pageHeight = 842.0;
        double marginLeft = options.MarginPoints;
        double marginRight = options.MarginPoints;
        double marginTop = options.MarginPoints;
        double marginBottom = options.MarginPoints;

        if (sectPr != null)
        {
            var pgSz = sectPr.GetFirstChild<PageSize>();
            if (pgSz?.Width?.Value != null && pgSz.Height?.Value != null)
            {
                double w = pgSz.Width.Value / 20.0;
                double h = pgSz.Height.Value / 20.0;
                if (pgSz.Orient?.Value == PageOrientationValues.Landscape)
                {
                    pageWidth = Math.Max(w, h);
                    pageHeight = Math.Min(w, h);
                }
                else
                {
                    pageWidth = Math.Min(w, h);
                    pageHeight = Math.Max(w, h);
                }
            }

            var pgMar = sectPr.GetFirstChild<PageMargin>();
            if (pgMar != null)
            {
                if (pgMar.Left?.Value != null) marginLeft = pgMar.Left.Value / 20.0;
                if (pgMar.Right?.Value != null) marginRight = pgMar.Right.Value / 20.0;
                if (pgMar.Top?.Value != null) marginTop = pgMar.Top.Value / 20.0;
                if (pgMar.Bottom?.Value != null) marginBottom = pgMar.Bottom.Value / 20.0;
            }
        }

        // Extract header and footer text if available
        string? headerText = null;
        string? footerText = null;
        if (sectPr != null)
        {
            var headerRef = sectPr.Elements<HeaderReference>().FirstOrDefault();
            if (headerRef?.Id?.Value != null && mainPart.TryGetPartById(headerRef.Id.Value, out var hPart) && hPart is HeaderPart headerPart)
            {
                headerText = headerPart.Header?.InnerText?.Trim();
            }

            var footerRef = sectPr.Elements<FooterReference>().FirstOrDefault();
            if (footerRef?.Id?.Value != null && mainPart.TryGetPartById(footerRef.Id.Value, out var fPart) && fPart is FooterPart footerPart)
            {
                footerText = footerPart.Footer?.InnerText?.Trim();
            }
        }

        using var pdf = new PdfDocument();
        pdf.Info.Title = Path.GetFileNameWithoutExtension(sourceDocxPath);
        pdf.Info.Creator = "SmartShift Document Engine (OpenXML to PDF)";
        pdf.Options.CompressContentStreams = true;

        double printableWidth = pageWidth - (marginLeft + marginRight);
        double printableHeight = pageHeight - (marginTop + marginBottom);

        var fontTitle = new XFont("Arial", 20, XFontStyleEx.Bold);
        var fontH1 = new XFont("Arial", 16, XFontStyleEx.Bold);
        var fontH2 = new XFont("Arial", 13, XFontStyleEx.Bold);
        var fontH3 = new XFont("Arial", 11.5, XFontStyleEx.Bold);
        var fontBody = new XFont("Arial", 10.5, XFontStyleEx.Regular);
        var fontBodyBold = new XFont("Arial", 10.5, XFontStyleEx.Bold);
        var fontBodyItalic = new XFont("Arial", 10.5, XFontStyleEx.Italic);
        var fontCaption = new XFont("Arial", 8.5, XFontStyleEx.Regular);

        var brushText = new XSolidBrush(XColor.FromArgb(20, 25, 35));
        var brushHeading = new XSolidBrush(XColor.FromArgb(15, 23, 42));
        var brushSubtext = new XSolidBrush(XColor.FromArgb(100, 116, 139));
        var penBorder = new XPen(XColor.FromArgb(203, 213, 225), 0.8);
        var penTable = new XPen(XColor.FromArgb(226, 232, 240), 0.6);
        var brushTableCell = new XSolidBrush(XColor.FromArgb(248, 250, 252));

        var pages = new List<PdfPage>();
        var gfxList = new List<XGraphics>();

        PdfPage CurrentPage()
        {
            if (pages.Count == 0)
            {
                var p = pdf.AddPage();
                p.Width = XUnit.FromPoint(pageWidth);
                p.Height = XUnit.FromPoint(pageHeight);
                pages.Add(p);
                gfxList.Add(XGraphics.FromPdfPage(p));
            }
            return pages[^1];
        }

        XGraphics CurrentGfx()
        {
            CurrentPage();
            return gfxList[^1];
        }

        double cursorY = marginTop;

        void AdvanceNewPage()
        {
            var newPage = pdf.AddPage();
            newPage.Width = XUnit.FromPoint(pageWidth);
            newPage.Height = XUnit.FromPoint(pageHeight);
            pages.Add(newPage);
            gfxList.Add(XGraphics.FromPdfPage(newPage));
            cursorY = marginTop;
        }

        void CheckNewPage(double neededHeight)
        {
            if (cursorY + neededHeight > pageHeight - marginBottom - 20)
            {
                AdvanceNewPage();
            }
        }

        var elements = body.ChildElements.ToList();
        int totalElements = elements.Count;

        for (int elemIdx = 0; elemIdx < totalElements; elemIdx++)
        {
            if (elemIdx % 10 == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int pct = 20 + (int)(65.0 * elemIdx / Math.Max(1, totalElements));
                progress?.Report(new ProcessingProgress(pct, $"Rendering document elements ({elemIdx + 1}/{totalElements})…"));
            }

            var element = elements[elemIdx];

            if (element is Paragraph para)
            {
                // Check explicit page break
                bool isPageBreak = para.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page) ||
                                   para.Descendants<LastRenderedPageBreak>().Any();

                if (isPageBreak && pages.Count > 0)
                {
                    AdvanceNewPage();
                }

                // Check if paragraph contains drawings / images
                var blips = para.Descendants<A.Blip>().ToList();
                if (blips.Count > 0)
                {
                    foreach (var blip in blips)
                    {
                        var embedId = blip.Embed?.Value;
                        if (!string.IsNullOrEmpty(embedId) && mainPart.TryGetPartById(embedId, out var part) && part is ImagePart imgPart)
                        {
                            try
                            {
                                using var imgStream = imgPart.GetStream(FileMode.Open, FileAccess.Read);
                                using var memStream = new MemoryStream();
                                imgStream.CopyTo(memStream);
                                memStream.Position = 0;

                                using var xImg = XImage.FromStream(memStream);
                                double maxImgWidth = Math.Min(printableWidth, 380);
                                double scale = Math.Min(1.0, maxImgWidth / xImg.PixelWidth);
                                double drawWidth = xImg.PixelWidth * scale;
                                double drawHeight = xImg.PixelHeight * scale;

                                if (drawHeight > printableHeight * 0.7)
                                {
                                    scale = (printableHeight * 0.7) / xImg.PixelHeight;
                                    drawWidth = xImg.PixelWidth * scale;
                                    drawHeight = xImg.PixelHeight * scale;
                                }

                                CheckNewPage(drawHeight + 16);
                                var gfx = CurrentGfx();
                                double imgX = marginLeft + ((printableWidth - drawWidth) / 2.0);
                                gfx.DrawImage(xImg, imgX, cursorY, drawWidth, drawHeight);
                                cursorY += drawHeight + 14;
                                imagesPreserved++;
                            }
                            catch
                            {
                                warnings.Add("An embedded image could not be decoded or rendered.");
                            }
                        }
                    }
                }

                var text = para.InnerText?.Trim();
                if (string.IsNullOrWhiteSpace(text))
                {
                    cursorY += 6;
                    continue;
                }

                totalChars += text.Length;

                var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                var (isHeading, level) = DetermineHeading(styleId);
                bool isBullet = para.ParagraphProperties?.NumberingProperties is not null;

                XFont currentFont;
                XBrush currentBrush;
                double lineHeight;
                double spaceBefore;
                double spaceAfter;
                double leftIndent = 0;

                if (isHeading)
                {
                    headingsDetected++;
                    if (level == 1)
                    {
                        currentFont = fontTitle;
                        currentBrush = brushHeading;
                        lineHeight = 24;
                        spaceBefore = 14;
                        spaceAfter = 8;
                    }
                    else if (level == 2)
                    {
                        currentFont = fontH1;
                        currentBrush = brushHeading;
                        lineHeight = 20;
                        spaceBefore = 12;
                        spaceAfter = 6;
                    }
                    else if (level == 3)
                    {
                        currentFont = fontH2;
                        currentBrush = brushHeading;
                        lineHeight = 17;
                        spaceBefore = 10;
                        spaceAfter = 4;
                    }
                    else
                    {
                        currentFont = fontH3;
                        currentBrush = brushHeading;
                        lineHeight = 15;
                        spaceBefore = 8;
                        spaceAfter = 4;
                    }
                }
                else if (isBullet)
                {
                    currentFont = fontBody;
                    currentBrush = brushText;
                    lineHeight = 14.5;
                    spaceBefore = 2;
                    spaceAfter = 3;
                    leftIndent = 16;
                }
                else
                {
                    currentFont = fontBody;
                    currentBrush = brushText;
                    lineHeight = 14.5;
                    spaceBefore = 3;
                    spaceAfter = 5;
                }

                cursorY += spaceBefore;

                var lines = WrapText(CurrentGfx(), text, currentFont, printableWidth - leftIndent);
                double blockHeight = lines.Count * lineHeight;
                CheckNewPage(blockHeight);

                var activeGfx = CurrentGfx();

                if (isBullet)
                {
                    activeGfx.DrawString("•", currentFont, currentBrush, marginLeft + 4, cursorY + (lineHeight * 0.75));
                }

                for (int l = 0; l < lines.Count; l++)
                {
                    CheckNewPage(lineHeight);
                    activeGfx = CurrentGfx();
                    activeGfx.DrawString(lines[l], currentFont, currentBrush, marginLeft + leftIndent, cursorY + (lineHeight * 0.75));
                    cursorY += lineHeight;
                }

                cursorY += spaceAfter;
            }
            else if (element is Table table)
            {
                var rows = table.Elements<TableRow>().ToList();
                if (rows.Count == 0) continue;

                int colCount = rows.Max(r => r.Elements<TableCell>().Count());
                if (colCount == 0) continue;

                tablesReconstructed++;
                double colWidth = printableWidth / colCount;
                cursorY += 8;

                bool isFirstRow = true;
                foreach (var row in rows)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var cells = row.Elements<TableCell>().ToList();
                    double rowHeight = 22;

                    foreach (var cell in cells)
                    {
                        var cellText = cell.InnerText?.Trim() ?? "";
                        var cellLines = WrapText(CurrentGfx(), cellText, fontBody, colWidth - 10);
                        double cellHeight = Math.Max(22, (cellLines.Count * 13) + 8);
                        if (cellHeight > rowHeight) rowHeight = cellHeight;
                    }

                    CheckNewPage(rowHeight);
                    var gfx = CurrentGfx();

                    for (int c = 0; c < colCount; c++)
                    {
                        double cellX = marginLeft + (c * colWidth);
                        var cellRect = new XRect(cellX, cursorY, colWidth, rowHeight);

                        if (isFirstRow)
                        {
                            gfx.DrawRectangle(brushTableCell, cellRect);
                        }
                        gfx.DrawRectangle(penTable, cellRect);

                        if (c < cells.Count)
                        {
                            var cellText = cells[c].InnerText?.Trim() ?? "";
                            totalChars += cellText.Length;
                            var cellFont = isFirstRow ? fontBodyBold : fontBody;
                            var cellLines = WrapText(gfx, cellText, cellFont, colWidth - 10);

                            double textY = cursorY + 6;
                            foreach (var line in cellLines)
                            {
                                gfx.DrawString(line, cellFont, brushText, cellX + 5, textY + 9);
                                textY += 13;
                            }
                        }
                    }

                    cursorY += rowHeight;
                    isFirstRow = false;
                }

                cursorY += 10;
            }
        }

        // Draw headers, footers, and page numbers across all pages
        if (pages.Count > 0)
        {
            for (int p = 0; p < pages.Count; p++)
            {
                var gfx = gfxList[p];

                // Header
                if (!string.IsNullOrWhiteSpace(headerText))
                {
                    gfx.DrawString(headerText, fontCaption, brushSubtext,
                        new XRect(marginLeft, marginTop - 18, printableWidth, 12),
                        XStringFormats.TopLeft);
                    gfx.DrawLine(penBorder, marginLeft, marginTop - 4, pageWidth - marginRight, marginTop - 4);
                }

                // Footer
                if (options.IncludePageNumbers || !string.IsNullOrWhiteSpace(footerText))
                {
                    gfx.DrawLine(penBorder, marginLeft, pageHeight - marginBottom + 8, pageWidth - marginRight, pageHeight - marginBottom + 8);
                    string displayFooter = !string.IsNullOrWhiteSpace(footerText)
                        ? $"{footerText} — Page {p + 1} of {pages.Count}"
                        : $"Page {p + 1} of {pages.Count}";

                    gfx.DrawString(displayFooter, fontCaption, brushSubtext,
                        new XRect(marginLeft, pageHeight - marginBottom + 12, printableWidth, 14),
                        XStringFormats.Center);
                }
            }
        }

        if (pages.Count == 0)
        {
            var p = pdf.AddPage();
            p.Width = XUnit.FromPoint(pageWidth);
            p.Height = XUnit.FromPoint(pageHeight);
            using var gfx = XGraphics.FromPdfPage(p);
            gfx.DrawString("(Empty Document)", fontBody, brushSubtext, marginLeft, marginTop + 20);
            pages.Add(p);
        }

        pagesPreserved = pages.Count;

        foreach (var gfx in gfxList)
        {
            gfx.Dispose();
        }

        progress?.Report(new ProcessingProgress(90, "Writing PDF document stream…"));
        pdf.Save(tempPartial);
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

    private static List<string> WrapText(XGraphics gfx, string text, XFont font, double maxWidth)
    {
        var lines = new List<string>();
        if (string.IsNullOrEmpty(text)) return lines;

        var words = text.Split([' '], StringSplitOptions.RemoveEmptyEntries);
        var currentLine = "";

        foreach (var word in words)
        {
            var testLine = string.IsNullOrEmpty(currentLine) ? word : $"{currentLine} {word}";
            var size = gfx.MeasureString(testLine, font);

            if (size.Width <= maxWidth)
            {
                currentLine = testLine;
            }
            else
            {
                if (!string.IsNullOrEmpty(currentLine))
                {
                    lines.Add(currentLine);
                }
                currentLine = word;
            }
        }

        if (!string.IsNullOrEmpty(currentLine))
        {
            lines.Add(currentLine);
        }

        return lines.Count > 0 ? lines : [text];
    }
}
