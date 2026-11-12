using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using SmartShift.Core.IO;
using SmartShift.Processing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace SmartShift.Processing.Documents;

public static class PptxToPdfConverter
{
    private const double EmuPerPoint = 12700.0;

    public static OfficeConversionResult Convert(
        string sourcePptxPath,
        string outputPath,
        string slideRange = "all",
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(sourcePptxPath) || !File.Exists(sourcePptxPath))
        {
            throw new FileNotFoundException("Source PowerPoint presentation does not exist.", sourcePptxPath);
        }

        var fileInfo = new FileInfo(sourcePptxPath);
        if (fileInfo.Length == 0)
        {
            throw new InvalidOperationException("Source PowerPoint presentation is empty (0 bytes).");
        }

        var sw = Stopwatch.StartNew();
        progress?.Report(new ProcessingProgress(10, "Opening PowerPoint presentation…"));

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

        int slidesPreserved = 0;
        int headingsDetected = 0;
        int shapesPreserved = 0;
        int imagesPreserved = 0;
        int tablesReconstructed = 0;
        long totalChars = 0;
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

                // Slide dimensions in points
                var slideSize = presentationPart.Presentation.SlideSize;
                double slideWidthPt = (slideSize?.Cx?.Value ?? 12192000L) / EmuPerPoint;
                double slideHeightPt = (slideSize?.Cy?.Value ?? 6858000L) / EmuPerPoint;
                if (slideWidthPt <= 0) slideWidthPt = 960.0;
                if (slideHeightPt <= 0) slideHeightPt = 540.0;

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

                using var pdf = new PdfDocument();
                pdf.Info.Title = Path.GetFileNameWithoutExtension(sourcePptxPath);
                pdf.Info.Creator = "SmartShift Advanced Office Conversion Engine (PPTX to PDF)";
                pdf.Options.CompressContentStreams = true;

                var fontTitle = new XFont("Arial", 22, XFontStyleEx.Bold);
                var fontSubtitle = new XFont("Arial", 16, XFontStyleEx.Regular);
                var fontBody = new XFont("Arial", 12, XFontStyleEx.Regular);
                var fontBodyBold = new XFont("Arial", 12, XFontStyleEx.Bold);
                var fontBodyItalic = new XFont("Arial", 12, XFontStyleEx.Italic);
                var fontSmall = new XFont("Arial", 9.5, XFontStyleEx.Regular);

                var brushDefaultText = new XSolidBrush(XColor.FromArgb(30, 41, 59));
                var brushTitleText = new XSolidBrush(XColor.FromArgb(15, 23, 42));
                var penShapeBorder = new XPen(XColor.FromArgb(203, 213, 225), 0.8);
                var penTableBorder = new XPen(XColor.FromArgb(226, 232, 240), 0.7);
                var brushTableCellBg = new XSolidBrush(XColor.FromArgb(248, 250, 252));

                for (int i = 0; i < selectedIndices.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int slideIdx = selectedIndices[i] - 1;
                    int pct = 15 + (int)(75.0 * i / selectedIndices.Count);
                    progress?.Report(new ProcessingProgress(pct, $"Rendering slide {i + 1} of {selectedIndices.Count}…"));

                    var slideId = slideIds[slideIdx];
                    var relId = slideId.RelationshipId?.Value;
                    if (string.IsNullOrEmpty(relId) || !presentationPart.TryGetPartById(relId, out var part) || part is not SlidePart slidePart)
                    {
                        warnings.Add($"Slide {selectedIndices[i]} could not be resolved from presentation package.");
                        continue;
                    }

                    var pdfPage = pdf.AddPage();
                    pdfPage.Width = XUnit.FromPoint(slideWidthPt);
                    pdfPage.Height = XUnit.FromPoint(slideHeightPt);
                    slidesPreserved++;

                    using var gfx = XGraphics.FromPdfPage(pdfPage);

                    // 1. Draw slide background
                    DrawSlideBackground(slidePart, gfx, slideWidthPt, slideHeightPt);

                    // 2. Process shapes tree
                    var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
                    if (shapeTree != null)
                    {
                        foreach (var element in shapeTree.ChildElements)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            if (element is P.Shape shape)
                            {
                                ProcessShape(shape, gfx, slideWidthPt, slideHeightPt,
                                    fontTitle, fontSubtitle, fontBody, fontBodyBold, fontBodyItalic,
                                    brushDefaultText, brushTitleText, penShapeBorder,
                                    ref headingsDetected, ref shapesPreserved, ref totalChars);
                            }
                            else if (element is P.Picture picture)
                            {
                                ProcessPicture(picture, slidePart, gfx, ref imagesPreserved);
                            }
                            else if (element is P.GraphicFrame frame)
                            {
                                ProcessGraphicFrame(frame, gfx, fontBody, fontBodyBold, brushDefaultText, penTableBorder, brushTableCellBg,
                                    ref tablesReconstructed, ref totalChars, limitations);
                            }
                            else if (element is P.GroupShape groupShape)
                            {
                                ProcessGroupShape(groupShape, slidePart, gfx, slideWidthPt, slideHeightPt,
                                    fontTitle, fontSubtitle, fontBody, fontBodyBold, fontBodyItalic,
                                    brushDefaultText, brushTitleText, penShapeBorder, penTableBorder, brushTableCellBg,
                                    ref headingsDetected, ref shapesPreserved, ref imagesPreserved, ref tablesReconstructed, ref totalChars, limitations);
                            }
                        }
                    }
                }

                if (pdf.PageCount == 0)
                {
                    var emptyPage = pdf.AddPage();
                    emptyPage.Width = XUnit.FromPoint(slideWidthPt);
                    emptyPage.Height = XUnit.FromPoint(slideHeightPt);
                    using var gfx = XGraphics.FromPdfPage(emptyPage);
                    gfx.DrawString("(No Slides Converted)", fontBody, brushDefaultText, 40, 60);
                }

                progress?.Report(new ProcessingProgress(92, "Writing PDF document stream…"));
                pdf.Save(tempPartial);
            }

            progress?.Report(new ProcessingProgress(96, "Validating generated PDF integrity…"));
            if (!File.Exists(tempPartial) || new FileInfo(tempPartial).Length == 0)
            {
                throw new InvalidOperationException("Failed to generate PDF output file.");
            }

            using (var verifyStream = new FileStream(tempPartial, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var verifyPdf = PdfReader.Open(verifyStream, PdfDocumentOpenMode.Import))
            {
                if (verifyPdf.PageCount <= 0)
                {
                    throw new InvalidOperationException("Generated PDF contains no pages.");
                }
            }

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
            File.Move(tempPartial, outputPath);
            sw.Stop();

            progress?.Report(new ProcessingProgress(100, "PowerPoint to PDF conversion completed."));

            var status = limitations.Count > 0 ? OfficeConversionStatus.SuccessfulWithLimitations : OfficeConversionStatus.Successful;
            var report = new OfficeConversionQualityReport(
                Status: status,
                ConversionMode: "PptxToPdfConversion",
                PagesOrSlidesCount: slidesPreserved,
                HeadingsDetectedCount: headingsDetected,
                TablesReconstructedCount: tablesReconstructed,
                ImagesPreservedCount: imagesPreserved,
                ShapesPreservedCount: shapesPreserved,
                CharacterCount: totalChars,
                Limitations: limitations,
                Warnings: warnings,
                ValidationPassed: true,
                Duration: sw.Elapsed);

            return new OfficeConversionResult(
                outputPath,
                report,
                $"Converted {slidesPreserved} presentation slide(s) to PDF with {shapesPreserved} shape(s) and {imagesPreserved} image(s).");
        }
        finally
        {
            if (File.Exists(tempPartial))
            {
                try { File.Delete(tempPartial); } catch { }
            }
        }
    }

    private static void DrawSlideBackground(SlidePart slidePart, XGraphics gfx, double slideWidth, double slideHeight)
    {
        // Default clean slide background
        gfx.DrawRectangle(new XSolidBrush(XColor.FromArgb(255, 255, 255)), 0, 0, slideWidth, slideHeight);

        var bg = slidePart.Slide?.CommonSlideData?.Background;
        var solidFill = bg?.Descendants<A.RgbColorModelHex>().FirstOrDefault();
        if (solidFill?.Val?.Value != null && solidFill.Val.Value.Length == 6)
        {
            try
            {
                var hex = solidFill.Val.Value;
                int r = System.Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = System.Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = System.Convert.ToInt32(hex.Substring(4, 2), 16);
                gfx.DrawRectangle(new XSolidBrush(XColor.FromArgb(r, g, b)), 0, 0, slideWidth, slideHeight);
            }
            catch { }
        }
    }

    private static void ProcessShape(
        P.Shape shape,
        XGraphics gfx,
        double slideWidth,
        double slideHeight,
        XFont fontTitle,
        XFont fontSubtitle,
        XFont fontBody,
        XFont fontBodyBold,
        XFont fontBodyItalic,
        XBrush brushDefaultText,
        XBrush brushTitleText,
        XPen penShapeBorder,
        ref int headingsDetected,
        ref int shapesPreserved,
        ref long totalChars)
    {
        var xfrm = shape.ShapeProperties?.Transform2D;
        if (xfrm == null) return;

        double x = (xfrm.Offset?.X?.Value ?? 0L) / EmuPerPoint;
        double y = (xfrm.Offset?.Y?.Value ?? 0L) / EmuPerPoint;
        double w = (xfrm.Extents?.Cx?.Value ?? 0L) / EmuPerPoint;
        double h = (xfrm.Extents?.Cy?.Value ?? 0L) / EmuPerPoint;

        if (w <= 0 || h <= 0) return;
        shapesPreserved++;

        // Draw background fill if shape has solidFill
        var solidFill = shape.ShapeProperties?.Descendants<A.RgbColorModelHex>().FirstOrDefault();
        if (solidFill?.Val?.Value != null && solidFill.Val.Value.Length == 6)
        {
            try
            {
                var hex = solidFill.Val.Value;
                int r = System.Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = System.Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = System.Convert.ToInt32(hex.Substring(4, 2), 16);
                gfx.DrawRectangle(new XSolidBrush(XColor.FromArgb(r, g, b)), x, y, w, h);
            }
            catch { }
        }

        // Draw outline if present
        if (shape.ShapeProperties?.Descendants<A.Outline>().Any() == true)
        {
            gfx.DrawRectangle(penShapeBorder, x, y, w, h);
        }

        var textBody = shape.TextBody;
        if (textBody == null) return;

        var paragraphs = textBody.Elements<A.Paragraph>().ToList();
        if (paragraphs.Count == 0) return;

        double currentY = y + 4;
        var phType = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value;
        bool isPlaceholderTitle = phType == P.PlaceholderValues.Title || phType == P.PlaceholderValues.CenteredTitle;

        if (isPlaceholderTitle)
        {
            headingsDetected++;
        }

        foreach (var para in paragraphs)
        {
            var runs = para.Elements<A.Run>().ToList();
            var paraText = string.Concat(runs.Select(r => r.Text?.Text ?? "")).Trim();
            if (string.IsNullOrWhiteSpace(paraText))
            {
                currentY += 12;
                continue;
            }

            totalChars += paraText.Length;

            // Determine font and size
            var firstRunPr = runs.FirstOrDefault()?.RunProperties;
            int fontSizeRaw = firstRunPr?.FontSize?.Value ?? 0;
            double fontSizePt = fontSizeRaw > 0 ? fontSizeRaw / 100.0 : (isPlaceholderTitle ? 22 : 12);

            bool isBold = firstRunPr?.Bold?.Value ?? isPlaceholderTitle;
            bool isItalic = firstRunPr?.Italic?.Value ?? false;

            XFont font;
            if (fontSizePt >= 20 || isPlaceholderTitle)
            {
                font = fontTitle;
                if (!isPlaceholderTitle) headingsDetected++;
            }
            else if (fontSizePt >= 15)
            {
                font = fontSubtitle;
            }
            else if (isBold)
            {
                font = fontBodyBold;
            }
            else if (isItalic)
            {
                font = fontBodyItalic;
            }
            else
            {
                font = fontBody;
            }

            var brush = isPlaceholderTitle ? brushTitleText : brushDefaultText;

            // Alignment
            var align = para.ParagraphProperties?.Alignment?.Value;
            var format = XStringFormats.TopLeft;
            double drawX = x + 4;
            if (align == A.TextAlignmentTypeValues.Center)
            {
                format = XStringFormats.TopCenter;
                drawX = x + (w / 2.0);
            }
            else if (align == A.TextAlignmentTypeValues.Right)
            {
                format = XStringFormats.TopRight;
                drawX = x + w - 4;
            }

            // Word wrap
            var lines = WrapText(gfx, paraText, font, Math.Max(20, w - 8));
            double lineHeight = Math.Max(12, fontSizePt * 1.25);

            foreach (var line in lines)
            {
                if (currentY + lineHeight > y + h + 10) break;
                gfx.DrawString(line, font, brush, drawX, currentY, format);
                currentY += lineHeight;
            }

            currentY += 4;
        }
    }

    private static void ProcessPicture(
        P.Picture picture,
        SlidePart slidePart,
        XGraphics gfx,
        ref int imagesPreserved)
    {
        var xfrm = picture.ShapeProperties?.Transform2D;
        if (xfrm == null) return;

        double x = (xfrm.Offset?.X?.Value ?? 0L) / EmuPerPoint;
        double y = (xfrm.Offset?.Y?.Value ?? 0L) / EmuPerPoint;
        double w = (xfrm.Extents?.Cx?.Value ?? 0L) / EmuPerPoint;
        double h = (xfrm.Extents?.Cy?.Value ?? 0L) / EmuPerPoint;

        if (w <= 0 || h <= 0) return;

        var embedId = picture.BlipFill?.Blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embedId) || !slidePart.TryGetPartById(embedId, out var part) || part is not ImagePart imagePart)
        {
            return;
        }

        try
        {
            using var stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            using var mem = new MemoryStream();
            stream.CopyTo(mem);
            mem.Position = 0;

            using var xImg = XImage.FromStream(mem);
            gfx.DrawImage(xImg, x, y, w, h);
            imagesPreserved++;
        }
        catch
        {
            // If image fails to decode, continue cleanly
        }
    }

    private static void ProcessGraphicFrame(
        P.GraphicFrame frame,
        XGraphics gfx,
        XFont fontBody,
        XFont fontBodyBold,
        XBrush brushDefaultText,
        XPen penTableBorder,
        XBrush brushTableCellBg,
        ref int tablesReconstructed,
        ref long totalChars,
        List<string> limitations)
    {
        var table = frame.Descendants<A.Table>().FirstOrDefault();
        if (table == null)
        {
            limitations.Add("Non-table graphic frame elements (charts or SmartArt) were skipped or rendered as outlines.");
            return;
        }

        var xfrm = frame.Transform;
        if (xfrm == null) return;

        double tableX = (xfrm.Offset?.X?.Value ?? 0L) / EmuPerPoint;
        double tableY = (xfrm.Offset?.Y?.Value ?? 0L) / EmuPerPoint;
        double tableW = (xfrm.Extents?.Cx?.Value ?? 0L) / EmuPerPoint;
        double tableH = (xfrm.Extents?.Cy?.Value ?? 0L) / EmuPerPoint;

        if (tableW <= 0 || tableH <= 0) return;

        var rows = table.Elements<A.TableRow>().ToList();
        if (rows.Count == 0) return;

        tablesReconstructed++;
        int colCount = rows.Max(r => r.Elements<A.TableCell>().Count());
        if (colCount == 0) return;

        double colWidth = tableW / colCount;
        double cursorY = tableY;

        bool isFirstRow = true;
        foreach (var row in rows)
        {
            double rowH = (row.Height?.Value ?? 0L) / EmuPerPoint;
            if (rowH <= 0) rowH = 22.0;

            var cells = row.Elements<A.TableCell>().ToList();
            for (int c = 0; c < colCount; c++)
            {
                double cellX = tableX + (c * colWidth);
                var cellRect = new XRect(cellX, cursorY, colWidth, rowH);

                if (isFirstRow)
                {
                    gfx.DrawRectangle(brushTableCellBg, cellRect);
                }
                gfx.DrawRectangle(penTableBorder, cellRect);

                if (c < cells.Count)
                {
                    var cellText = cells[c].TextBody?.InnerText?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(cellText))
                    {
                        totalChars += cellText.Length;
                        var cellFont = isFirstRow ? fontBodyBold : fontBody;
                        var lines = WrapText(gfx, cellText, cellFont, colWidth - 8);
                        double textY = cursorY + 4;
                        foreach (var line in lines)
                        {
                            if (textY + 12 > cursorY + rowH) break;
                            gfx.DrawString(line, cellFont, brushDefaultText, cellX + 4, textY + 9);
                            textY += 12;
                        }
                    }
                }
            }

            cursorY += rowH;
            isFirstRow = false;
        }
    }

    private static void ProcessGroupShape(
        P.GroupShape groupShape,
        SlidePart slidePart,
        XGraphics gfx,
        double slideWidth,
        double slideHeight,
        XFont fontTitle,
        XFont fontSubtitle,
        XFont fontBody,
        XFont fontBodyBold,
        XFont fontBodyItalic,
        XBrush brushDefaultText,
        XBrush brushTitleText,
        XPen penShapeBorder,
        XPen penTableBorder,
        XBrush brushTableCellBg,
        ref int headingsDetected,
        ref int shapesPreserved,
        ref int imagesPreserved,
        ref int tablesReconstructed,
        ref long totalChars,
        List<string> limitations)
    {
        foreach (var child in groupShape.ChildElements)
        {
            if (child is P.Shape shape)
            {
                ProcessShape(shape, gfx, slideWidth, slideHeight,
                    fontTitle, fontSubtitle, fontBody, fontBodyBold, fontBodyItalic,
                    brushDefaultText, brushTitleText, penShapeBorder,
                    ref headingsDetected, ref shapesPreserved, ref totalChars);
            }
            else if (child is P.Picture pic)
            {
                ProcessPicture(pic, slidePart, gfx, ref imagesPreserved);
            }
            else if (child is P.GraphicFrame frame)
            {
                ProcessGraphicFrame(frame, gfx, fontBody, fontBodyBold, brushDefaultText, penTableBorder, brushTableCellBg,
                    ref tablesReconstructed, ref totalChars, limitations);
            }
        }
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
