using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf.IO;
using SkiaSharp;
using SmartShift.Core.Actions;
using SmartShift.Core.Capabilities;
using SmartShift.Core.Classification;
using SmartShift.Core.History;
using SmartShift.Core.IO;
using SmartShift.Infrastructure.Output;
using SmartShift.Processing.Conversion;
using SmartShift.Processing.Documents;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace SmartShift.Processing.Tests;

public sealed class AdvancedOfficeConversionEngineTests : IDisposable
{
    private readonly string _tempDir;
    private readonly FileClassifier _classifier;
    private readonly OutputPathAllocator _allocator;
    private readonly OfficeConversionEngine _engine;

    public AdvancedOfficeConversionEngineTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "smartshift-phase3-step2-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
        _classifier = new FileClassifier();
        _allocator = new OutputPathAllocator(_tempDir);
        _engine = new OfficeConversionEngine(_allocator);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch { }
    }

    private static string ComputeSha256(string path)
    {
        using var stream = File.OpenRead(path);
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(stream));
    }

    // ==========================================
    // TEST FIXTURES GENERATORS
    // ==========================================

    private string CreateSimpleDocx()
    {
        var path = Path.Combine(_tempDir, $"doc_{Guid.NewGuid():N}.docx");
        using (var wordDoc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new W.Document(
                new W.Body(
                    new W.Paragraph(
                        new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Heading1" }),
                        new W.Run(new W.Text("Project Architecture Overview"))),
                    new W.Paragraph(
                        new W.Run(new W.RunProperties(new W.Bold()), new W.Text("Executive Summary: ")),
                        new W.Run(new W.Text("SmartShift transforms enterprise desktop document processing."))),
                    new W.Paragraph(
                        new W.ParagraphProperties(new W.NumberingProperties()),
                        new W.Run(new W.Text("Feature 1: Local-first offline execution"))),
                    new W.Paragraph(
                        new W.ParagraphProperties(new W.NumberingProperties()),
                        new W.Run(new W.Text("Feature 2: Strict source immutability Unicode ★ € £")))));
            mainPart.Document.Save();
        }
        return path;
    }

    private string CreateMultiPageDocx()
    {
        var path = Path.Combine(_tempDir, $"multipage_{Guid.NewGuid():N}.docx");
        using (var wordDoc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();

            // Add Header and Footer parts
            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader1");
            headerPart.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text("SmartShift Confidential Report"))));
            headerPart.Header.Save();

            var footerPart = mainPart.AddNewPart<FooterPart>("rIdFooter1");
            footerPart.Footer = new W.Footer(new W.Paragraph(new W.Run(new W.Text("Internal Distribution Only"))));
            footerPart.Footer.Save();

            var body = new W.Body();

            // Page 1 content
            body.AppendChild(new W.Paragraph(
                new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Heading1" }),
                new W.Run(new W.Text("Section 1: Introduction"))));
            body.AppendChild(new W.Paragraph(new W.Run(new W.Text("First page narrative details."))));

            // Explicit page break
            body.AppendChild(new W.Paragraph(new W.Run(new W.Break { Type = W.BreakValues.Page })));

            // Page 2 content
            body.AppendChild(new W.Paragraph(
                new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Heading2" }),
                new W.Run(new W.Text("Section 2: Landscape Details"))));
            body.AppendChild(new W.Paragraph(new W.Run(new W.Text("Second page content across page boundary."))));

            // Section properties with landscape orientation and headers/footers
            var sectPr = new W.SectionProperties(
                new W.HeaderReference { Id = "rIdHeader1", Type = W.HeaderFooterValues.Default },
                new W.FooterReference { Id = "rIdFooter1", Type = W.HeaderFooterValues.Default },
                new W.PageSize { Width = 16840U, Height = 11900U, Orient = W.PageOrientationValues.Landscape },
                new W.PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 });
            body.AppendChild(sectPr);

            mainPart.Document = new W.Document(body);
            mainPart.Document.Save();
        }
        return path;
    }

    private string CreateTableDocx()
    {
        var path = Path.Combine(_tempDir, $"table_{Guid.NewGuid():N}.docx");
        using (var wordDoc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            var body = new W.Body(
                new W.Paragraph(
                    new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Title" }),
                    new W.Run(new W.Text("Quarterly Metrics"))),
                new W.Table(
                    new W.TableRow(
                        new W.TableCell(new W.Paragraph(new W.Run(new W.Text("Period")))),
                        new W.TableCell(new W.Paragraph(new W.Run(new W.Text("Revenue")))),
                        new W.TableCell(new W.Paragraph(new W.Run(new W.Text("Growth"))))),
                    new W.TableRow(
                        new W.TableCell(new W.Paragraph(new W.Run(new W.Text("Q1")))),
                        new W.TableCell(new W.Paragraph(new W.Run(new W.Text("$1,500,000")))),
                        new W.TableCell(new W.Paragraph(new W.Run(new W.Text("+15%"))))),
                    new W.TableRow(
                        new W.TableCell(new W.Paragraph(new W.Run(new W.Text("Q2")))),
                        new W.TableCell(new W.Paragraph(new W.Run(new W.Text("$1,850,000")))),
                        new W.TableCell(new W.Paragraph(new W.Run(new W.Text("+23%")))))));

            mainPart.Document = new W.Document(body);
            mainPart.Document.Save();
        }
        return path;
    }

    private string CreateImageDocx()
    {
        var path = Path.Combine(_tempDir, $"image_{Guid.NewGuid():N}.docx");
        using (var wordDoc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            var imgPart = mainPart.AddImagePart(ImagePartType.Png);

            // Generate small 32x32 test PNG
            using var bmp = new SKBitmap(32, 32);
            using (var canvas = new SKCanvas(bmp))
            {
                canvas.Clear(SKColors.Blue);
            }
            using var data = bmp.Encode(SKEncodedImageFormat.Png, 100);
            using (var stream = imgPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                data.SaveTo(stream);
            }

            var relId = mainPart.GetIdOfPart(imgPart);

            var drawing = new W.Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 990000L, Cy = 792000L },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties { Id = 1U, Name = "Picture 1" },
                    new A.Graphic(
                        new A.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties { Id = 1U, Name = "Image" },
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                    new A.Blip { Embed = relId }),
                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                    new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 990000L, Cy = 792000L }),
                                    new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle })))
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));

            var body = new W.Body(
                new W.Paragraph(new W.Run(new W.Text("Document with Embedded Graphic"))),
                new W.Paragraph(new W.Run(drawing)));

            mainPart.Document = new W.Document(body);
            mainPart.Document.Save();
        }
        return path;
    }

    private string CreateMultiSlidePptx()
    {
        var path = Path.Combine(_tempDir, $"pres_{Guid.NewGuid():N}.pptx");
        using (var pptxDoc = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presPart = pptxDoc.AddPresentationPart();
            presPart.Presentation = new P.Presentation();

            // Slide size: 16:9 widescreen (960 x 540 pt)
            presPart.Presentation.SlideSize = new P.SlideSize { Cx = 12192000, Cy = 6858000, Type = P.SlideSizeValues.Screen16x9 };
            var slideIdList = new P.SlideIdList();
            presPart.Presentation.SlideIdList = slideIdList;

            // Slide 1: Title & Subtitle
            var slidePart1 = presPart.AddNewPart<SlidePart>("rIdSlide1");
            var shapeTree1 = new P.ShapeTree();
            shapeTree1.AppendChild(new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()));
            shapeTree1.AppendChild(new P.GroupShapeProperties(new A.TransformGroup()));

            // Title Shape
            var titleShape = new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 2, Name = "Title 1" },
                    new P.NonVisualShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
                new P.ShapeProperties(new A.Transform2D(new A.Offset { X = 1270000, Y = 635000 }, new A.Extents { Cx = 9652000, Cy = 1270000 })),
                new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.RunProperties { FontSize = 2400, Bold = true }, new A.Text("SmartShift Office Platform")))));
            shapeTree1.AppendChild(titleShape);

            slidePart1.Slide = new P.Slide(new P.CommonSlideData(shapeTree1));
            slidePart1.Slide.Save();
            slideIdList.AppendChild(new P.SlideId { Id = 256, RelationshipId = "rIdSlide1" });

            // Slide 2: Table & Bullets
            var slidePart2 = presPart.AddNewPart<SlidePart>("rIdSlide2");
            var shapeTree2 = new P.ShapeTree();
            shapeTree2.AppendChild(new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()));
            shapeTree2.AppendChild(new P.GroupShapeProperties(new A.TransformGroup()));

            var slide2Title = new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 3, Name = "Title 2" },
                    new P.NonVisualShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
                new P.ShapeProperties(new A.Transform2D(new A.Offset { X = 635000, Y = 381000 }, new A.Extents { Cx = 8890000, Cy = 889000 })),
                new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.RunProperties { FontSize = 2000, Bold = true }, new A.Text("Performance Overview")))));
            shapeTree2.AppendChild(slide2Title);

            // Table GraphicFrame
            var tableFrame = new P.GraphicFrame(
                new P.NonVisualGraphicFrameProperties(
                    new P.NonVisualDrawingProperties { Id = 4, Name = "Table 1" },
                    new P.NonVisualGraphicFrameDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.Transform(new A.Offset { X = 635000, Y = 1524000 }, new A.Extents { Cx = 8890000, Cy = 2540000 }),
                new A.Graphic(
                    new A.GraphicData(
                        new A.Table(
                            new A.TableProperties(),
                            new A.TableGrid(new A.GridColumn { Width = 4445000 }, new A.GridColumn { Width = 4445000 }),
                            new A.TableRow(
                                new A.TableCell(new A.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("Engine"))))),
                                new A.TableCell(new A.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("Throughput")))))),
                            new A.TableRow(
                                new A.TableCell(new A.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("Office Converter"))))),
                                new A.TableCell(new A.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("120 fps"))))))))
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
                    }));
            shapeTree2.AppendChild(tableFrame);

            slidePart2.Slide = new P.Slide(new P.CommonSlideData(shapeTree2));
            slidePart2.Slide.Save();
            slideIdList.AppendChild(new P.SlideId { Id = 257, RelationshipId = "rIdSlide2" });

            presPart.Presentation.Save();
        }
        return path;
    }

    // ==========================================
    // 1. DOCX-TO-PDF TESTS
    // ==========================================

    [Fact]
    public async Task ConvertDocxToPdf_SinglePage_PreservesContentAndHeadings()
    {
        var docxPath = CreateSimpleDocx();
        var shaBefore = ComputeSha256(docxPath);

        var result = await _engine.ConvertDocxToPdfAsync(docxPath, null, null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(docxPath), shaBefore);
        Assert.NotNull(result.QualityReport);
        Assert.Equal(OfficeConversionStatus.Successful, result.QualityReport.Status);
        Assert.True(result.QualityReport.HeadingsDetectedCount >= 1);
        Assert.True(result.QualityReport.CharacterCount > 20);

        // Verify PDF integrity
        using var stream = File.OpenRead(result.OutputPath);
        using var pdfDoc = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
        Assert.True(pdfDoc.PageCount >= 1);
    }

    [Fact]
    public async Task ConvertDocxToPdf_MultiPage_PreservesPageBreaksOrientationAndFooters()
    {
        var docxPath = CreateMultiPageDocx();
        var shaBefore = ComputeSha256(docxPath);

        var result = await _engine.ConvertDocxToPdfAsync(docxPath, null, null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(docxPath), shaBefore);
        Assert.NotNull(result.QualityReport);
        Assert.Equal(OfficeConversionStatus.Successful, result.QualityReport.Status);

        using var stream = File.OpenRead(result.OutputPath);
        using var pdfDoc = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
        // Multi-page document with explicit page break must generate >= 2 pages
        Assert.True(pdfDoc.PageCount >= 2);
    }

    [Fact]
    public async Task ConvertDocxToPdf_TableDocx_PreservesTableStructure()
    {
        var docxPath = CreateTableDocx();
        var shaBefore = ComputeSha256(docxPath);

        var result = await _engine.ConvertDocxToPdfAsync(docxPath, null, null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(docxPath), shaBefore);
        Assert.NotNull(result.QualityReport);
        Assert.True(result.QualityReport.TablesReconstructedCount >= 1);

        using var stream = File.OpenRead(result.OutputPath);
        using var pdfDoc = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
        Assert.True(pdfDoc.PageCount >= 1);
    }

    [Fact]
    public async Task ConvertDocxToPdf_ImageDocx_PreservesImagePlacement()
    {
        var docxPath = CreateImageDocx();
        var shaBefore = ComputeSha256(docxPath);

        var result = await _engine.ConvertDocxToPdfAsync(docxPath, null, null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(docxPath), shaBefore);
        Assert.NotNull(result.QualityReport);
        Assert.True(result.QualityReport.ImagesPreservedCount >= 1);

        using var stream = File.OpenRead(result.OutputPath);
        using var pdfDoc = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
        Assert.True(pdfDoc.PageCount >= 1);
    }

    [Fact]
    public async Task ConvertDocxToPdf_InvalidOrEmpty_ThrowsGracefully()
    {
        var emptyDocx = Path.Combine(_tempDir, "empty.docx");
        File.WriteAllBytes(emptyDocx, []);

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            _engine.ConvertDocxToPdfAsync(emptyDocx));

        await Assert.ThrowsAsync<FileNotFoundException>(() =>
            _engine.ConvertDocxToPdfAsync(Path.Combine(_tempDir, "missing.docx")));
    }

    // ==========================================
    // 2. PPTX-TO-PDF TESTS
    // ==========================================

    [Fact]
    public async Task ConvertPptxToPdf_MultiSlide_PreservesSlidesShapesAndOrder()
    {
        var pptxPath = CreateMultiSlidePptx();
        var shaBefore = ComputeSha256(pptxPath);

        var result = await _engine.ConvertPptxToPdfAsync(pptxPath, "all", null, null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(pptxPath), shaBefore);
        Assert.NotNull(result.QualityReport);
        Assert.Equal(OfficeConversionStatus.Successful, result.QualityReport.Status);
        Assert.Equal(2, result.QualityReport.PagesOrSlidesCount);
        Assert.True(result.QualityReport.HeadingsDetectedCount >= 1);
        Assert.True(result.QualityReport.TablesReconstructedCount >= 1);

        using var stream = File.OpenRead(result.OutputPath);
        using var pdfDoc = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
        Assert.Equal(2, pdfDoc.PageCount);
    }

    [Fact]
    public async Task ConvertPptxToPdf_SlideRange_ExportsSelectedOnly()
    {
        var pptxPath = CreateMultiSlidePptx();
        var shaBefore = ComputeSha256(pptxPath);

        var result = await _engine.ConvertPptxToPdfAsync(pptxPath, "1", null, null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(pptxPath), shaBefore);
        Assert.NotNull(result.QualityReport);
        Assert.Equal(1, result.QualityReport.PagesOrSlidesCount);

        using var stream = File.OpenRead(result.OutputPath);
        using var pdfDoc = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
        Assert.Equal(1, pdfDoc.PageCount);
    }

    // ==========================================
    // 3. PPTX-TO-IMAGES TESTS (PNG, JPEG, WebP)
    // ==========================================

    [Theory]
    [InlineData(SlideExportFormat.Png, ".png", 100)]
    [InlineData(SlideExportFormat.Jpeg, ".jpg", 150)]
    [InlineData(SlideExportFormat.Webp, ".webp", 72)]
    public async Task ExportPptxSlides_FormatsAndResolutions_GeneratesValidImages(SlideExportFormat format, string expectedExt, int dpi)
    {
        var pptxPath = CreateMultiSlidePptx();
        var shaBefore = ComputeSha256(pptxPath);

        var result = await _engine.ExportPptxSlidesAsync(pptxPath, format, dpi, 85, "all", null, null, CancellationToken.None);

        Assert.NotEmpty(result.OutputPaths);
        Assert.Equal(2, result.OutputPaths.Count);
        Assert.Equal(ComputeSha256(pptxPath), shaBefore);
        Assert.Equal(OfficeConversionStatus.Successful, result.QualityReport.Status);

        foreach (var file in result.OutputPaths)
        {
            Assert.True(File.Exists(file));
            Assert.EndsWith(expectedExt, file, StringComparison.OrdinalIgnoreCase);

            using var codec = SKCodec.Create(file);
            Assert.NotNull(codec);
            Assert.True(codec.Info.Width > 0);
            Assert.True(codec.Info.Height > 0);
        }
    }

    [Fact]
    public async Task ExportPptxSlides_Cancellation_CleansUpStagedFiles()
    {
        var pptxPath = CreateMultiSlidePptx();
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            _engine.ExportPptxSlidesAsync(pptxPath, SlideExportFormat.Png, 150, 85, "all", null, null, cts.Token));

        var files = Directory.GetFiles(_tempDir, "*.png");
        Assert.Empty(files);
    }

    // ==========================================
    // 4. DOCX-TO-TXT TESTS
    // ==========================================

    [Fact]
    public async Task ExtractDocxText_ParagraphsHeadingsTables_PreservesDocumentOrder()
    {
        var docxPath = CreateTableDocx();
        var shaBefore = ComputeSha256(docxPath);

        var result = await _engine.ExtractDocxTextAsync(docxPath, null, null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(docxPath), shaBefore);
        Assert.NotNull(result.QualityReport);
        Assert.Equal(OfficeConversionStatus.Successful, result.QualityReport.Status);

        var text = File.ReadAllText(result.OutputPath, Encoding.UTF8);
        Assert.Contains("Quarterly Metrics", text);
        Assert.Contains("Revenue", text);
        Assert.Contains("$1,500,000", text);
    }

    [Fact]
    public async Task ExtractDocxText_PageBreaks_EmitsExplicitPageBreakMarkers()
    {
        var docxPath = CreateMultiPageDocx();
        var shaBefore = ComputeSha256(docxPath);

        var result = await _engine.ExtractDocxTextAsync(docxPath, null, null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(docxPath), shaBefore);

        var text = File.ReadAllText(result.OutputPath, Encoding.UTF8);
        Assert.Contains("--- Page Break ---", text);
        Assert.Contains("Section 1: Introduction", text);
        Assert.Contains("Section 2: Landscape Details", text);
    }

    // ==========================================
    // 5. PPTX-TO-TXT TESTS
    // ==========================================

    [Fact]
    public async Task ExtractPptxText_SlideOrderTitlesAndTables_PreservesStructure()
    {
        var pptxPath = CreateMultiSlidePptx();
        var shaBefore = ComputeSha256(pptxPath);

        var result = await _engine.ExtractPptxTextAsync(pptxPath, "all", null, null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(pptxPath), shaBefore);
        Assert.NotNull(result.QualityReport);
        Assert.Equal(2, result.QualityReport.PagesOrSlidesCount);

        var text = File.ReadAllText(result.OutputPath, Encoding.UTF8);
        Assert.Contains("--- Slide 1: SmartShift Office Platform ---", text);
        Assert.Contains("--- Slide 2: Performance Overview ---", text);
        Assert.Contains("Office Converter", text);
        Assert.Contains("120 fps", text);
    }

    // ==========================================
    // 6. DOCX-TO-HTML TESTS
    // ==========================================

    [Fact]
    public async Task ConvertDocxToHtml_ValidHtml5_HeadingsParagraphsListsAndTables()
    {
        var docxPath = CreateTableDocx();
        var shaBefore = ComputeSha256(docxPath);

        var result = await _engine.ConvertDocxToHtmlAsync(docxPath, null, null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(docxPath), shaBefore);
        Assert.NotNull(result.QualityReport);

        var html = File.ReadAllText(result.OutputPath, Encoding.UTF8);
        Assert.Contains("<!DOCTYPE html>", html);
        Assert.Contains("<h1>Quarterly Metrics</h1>", html);
        Assert.Contains("<table>", html);
        Assert.Contains("<th>Period</th>", html);
        Assert.Contains("<td>$1,500,000</td>", html);
    }

    [Fact]
    public async Task ConvertDocxToHtml_EmbeddedImages_GeneratesSelfContainedDataUris()
    {
        var docxPath = CreateImageDocx();
        var shaBefore = ComputeSha256(docxPath);

        var result = await _engine.ConvertDocxToHtmlAsync(docxPath, null, null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(docxPath), shaBefore);
        Assert.NotNull(result.QualityReport);
        Assert.True(result.QualityReport.ImagesPreservedCount >= 1);

        var html = File.ReadAllText(result.OutputPath, Encoding.UTF8);
        Assert.Contains("<img src=\"data:image/", html);
        Assert.Contains(";base64,", html);
    }

    // ==========================================
    // 7. INTEGRATION, DI & PROCESSOR CATALOG TESTS
    // ==========================================

    [Fact]
    public void ProcessorCatalog_RegistersOfficeProcessors_Successfully()
    {
        var catalog = new ProcessorCatalog(_classifier, _allocator);

        Assert.NotNull(catalog.Find(ActionIds.ConvertDocxToPdf));
        Assert.NotNull(catalog.Find(ActionIds.ConvertDocxToTxt));
        Assert.NotNull(catalog.Find(ActionIds.ConvertDocxToHtml));
        Assert.NotNull(catalog.Find(ActionIds.ConvertPptxToPdf));
        Assert.NotNull(catalog.Find(ActionIds.ExportPptxImages));
        Assert.NotNull(catalog.Find(ActionIds.ExtractPptxText));
    }

    [Fact]
    public async Task UnifiedConversion_DocxAndPptx_RoutesAndExecutesWithHistory()
    {
        var capabilities = new FileCapabilityRegistry();
        var catalog = new ProcessorCatalog(_classifier, _allocator);
        var unified = new UnifiedConversionService(capabilities, catalog, _allocator, _classifier);

        var docxPath = CreateSimpleDocx();
        var classifiedDocx = _classifier.Classify(docxPath);

        // Convert DOCX to TXT via UnifiedConvert
        var result = await unified.ConvertAsync(
            classifiedDocx,
            DetectedFormat.Txt,
            new Dictionary<string, string>(),
            _tempDir,
            null,
            CancellationToken.None);

        Assert.True(result.Succeeded);
        Assert.True(File.Exists(result.OutputPath));

        var text = File.ReadAllText(result.OutputPath, Encoding.UTF8);
        Assert.Contains("Project Architecture Overview", text);
    }
}
