using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using SkiaSharp;
using SmartShift.Core.Actions;
using SmartShift.Core.Classification;
using SmartShift.Core.History;
using SmartShift.Core.IO;
using SmartShift.Infrastructure.Output;
using SmartShift.Processing.Pdf;
using SmartShift.Processing.Pdf.Office;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace SmartShift.Processing.Tests;

public sealed class AdvancedPdfConversionEngineTests : IDisposable
{
    private readonly string _tempDir;
    private readonly FileClassifier _classifier;
    private readonly OutputPathAllocator _allocator;
    private readonly PdfConversionEngine _engine;

    public AdvancedPdfConversionEngineTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "smartshift-phase3-step1-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
        _classifier = new FileClassifier();
        _allocator = new OutputPathAllocator(_tempDir);
        _engine = new PdfConversionEngine(_allocator);
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

    private string CreateTwoColumnPdf()
    {
        var path = Path.Combine(_tempDir, $"twocol_{Guid.NewGuid():N}.pdf");
        using (var doc = new PdfDocument())
        {
            var page = doc.AddPage();
            page.Width = XUnit.FromPoint(600);
            page.Height = XUnit.FromPoint(800);

            using var gfx = XGraphics.FromPdfPage(page);
            var headerFont = new XFont("Arial", 18);
            var colFont = new XFont("Arial", 11);

            // Document Header
            gfx.DrawString("Executive Annual Summary", headerFont, XBrushes.DarkSlateBlue, new XPoint(50, 60));

            // Column 1 (Left: X=50)
            gfx.DrawString("Left column paragraph line 1.", colFont, XBrushes.Black, new XPoint(50, 110));
            gfx.DrawString("Left column paragraph line 2.", colFont, XBrushes.Black, new XPoint(50, 130));

            // Column 2 (Right: X=320)
            gfx.DrawString("Right column paragraph line 1.", colFont, XBrushes.Black, new XPoint(320, 110));
            gfx.DrawString("Right column paragraph line 2.", colFont, XBrushes.Black, new XPoint(320, 130));

            doc.Save(path);
        }
        return path;
    }

    private string CreateTablePdf()
    {
        var path = Path.Combine(_tempDir, $"table_{Guid.NewGuid():N}.pdf");
        using (var doc = new PdfDocument())
        {
            var page = doc.AddPage();
            page.Width = XUnit.FromPoint(600);
            page.Height = XUnit.FromPoint(800);

            using var gfx = XGraphics.FromPdfPage(page);
            var titleFont = new XFont("Arial", 16);
            var cellFont = new XFont("Arial", 10);

            gfx.DrawString("Financial Balance Sheet", titleFont, XBrushes.Black, new XPoint(50, 50));

            // Table header row (3 columns with wide horizontal gap > 20 pt)
            gfx.DrawString("Quarter", cellFont, XBrushes.Black, new XPoint(50, 100));
            gfx.DrawString("Revenue", cellFont, XBrushes.Black, new XPoint(200, 100));
            gfx.DrawString("NetProfit", cellFont, XBrushes.Black, new XPoint(350, 100));

            // Table row 1
            gfx.DrawString("Q1-2026", cellFont, XBrushes.Black, new XPoint(50, 125));
            gfx.DrawString("$1,200,000", cellFont, XBrushes.Black, new XPoint(200, 125));
            gfx.DrawString("$340,000", cellFont, XBrushes.Black, new XPoint(350, 125));

            // Table row 2
            gfx.DrawString("Q2-2026", cellFont, XBrushes.Black, new XPoint(50, 150));
            gfx.DrawString("$1,450,000", cellFont, XBrushes.Black, new XPoint(200, 150));
            gfx.DrawString("$410,000", cellFont, XBrushes.Black, new XPoint(350, 150));

            doc.Save(path);
        }
        return path;
    }

    private string CreateMixedTextAndImagePdf()
    {
        var path = Path.Combine(_tempDir, $"mixed_{Guid.NewGuid():N}.pdf");

        // Create a small in-memory PNG
        using var bmp = new SKBitmap(100, 100);
        using (var canvas = new SKCanvas(bmp))
        {
            canvas.Clear(SKColors.CornflowerBlue);
        }
        using var imgData = bmp.Encode(SKEncodedImageFormat.Png, 100);
        var imgPath = Path.Combine(_tempDir, $"img_{Guid.NewGuid():N}.png");
        File.WriteAllBytes(imgPath, imgData.ToArray());

        using (var doc = new PdfDocument())
        {
            var page = doc.AddPage();
            page.Width = XUnit.FromPoint(600);
            page.Height = XUnit.FromPoint(800);

            using var gfx = XGraphics.FromPdfPage(page);
            var font = new XFont("Arial", 12);
            gfx.DrawString("Heading with Illustration", font, XBrushes.Black, new XPoint(50, 50));

            using var ximg = XImage.FromFile(imgPath);
            gfx.DrawImage(ximg, 50, 80, 200, 150);

            gfx.DrawString("Caption below figure", font, XBrushes.DarkGray, new XPoint(50, 250));

            doc.Save(path);
        }
        try { File.Delete(imgPath); } catch { }
        return path;
    }

    private string CreateEmptyPagePdf()
    {
        var path = Path.Combine(_tempDir, $"empty_{Guid.NewGuid():N}.pdf");
        using (var doc = new PdfDocument())
        {
            doc.AddPage();
            doc.Save(path);
        }
        return path;
    }

    // ==========================================
    // 1. PDF-to-DOCX CONVERSION TESTS
    // ==========================================

    [Fact]
    public async Task ConvertToDocxAsync_TwoColumnPdf_ExtractsInReadingOrder()
    {
        var pdfPath = CreateTwoColumnPdf();
        var shaBefore = ComputeSha256(pdfPath);

        var result = await _engine.ConvertToDocxAsync(pdfPath, "all", null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(pdfPath), shaBefore); // Immutability
        Assert.NotNull(result.QualityReport);
        Assert.Equal(PdfConversionStatus.Successful, result.QualityReport.Status);
        Assert.True(result.QualityReport.HeadingsDetectedCount >= 1);

        // Verify OpenXML structure and reading order
        var validator = new OpenXmlValidator();
        using (var docx = WordprocessingDocument.Open(result.OutputPath, false))
        {
            var validationErrors = validator.Validate(docx);
            Assert.DoesNotContain(validationErrors, e => e.ErrorType == ValidationErrorType.Schema && !e.Description.Contains("attribute"));

            Assert.NotNull(docx.MainDocumentPart?.Document?.Body);
            var bodyText = docx.MainDocumentPart.Document.Body.InnerText;
            Assert.Contains("Executive Annual Summary", bodyText);
            Assert.Contains("Left column paragraph", bodyText);
            Assert.Contains("Right column paragraph", bodyText);

            // Left column should appear before right column in natural reading order
            var leftIdx = bodyText.IndexOf("Left column paragraph line 1", StringComparison.Ordinal);
            var rightIdx = bodyText.IndexOf("Right column paragraph line 1", StringComparison.Ordinal);
            Assert.True(leftIdx < rightIdx, $"Expected leftIdx ({leftIdx}) < rightIdx ({rightIdx}). Body: [{bodyText}]");
        }
    }

    [Fact]
    public async Task ConvertToDocxAsync_TablePdf_ReconstructsOpenXmlTable()
    {
        var pdfPath = CreateTablePdf();
        var shaBefore = ComputeSha256(pdfPath);

        var result = await _engine.ConvertToDocxAsync(pdfPath, "all", null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(pdfPath), shaBefore);
        Assert.NotNull(result.QualityReport);
        Assert.True(result.QualityReport.TablesReconstructedCount >= 1);

        using (var docx = WordprocessingDocument.Open(result.OutputPath, false))
        {
            Assert.NotNull(docx.MainDocumentPart?.Document?.Body);
            var tables = docx.MainDocumentPart.Document.Body.Descendants<Table>().ToList();
            Assert.NotEmpty(tables);

            var firstTable = tables[0];
            var rows = firstTable.Descendants<TableRow>().ToList();
            Assert.True(rows.Count >= 2);

            var tableText = firstTable.InnerText;
            Assert.Contains("Quarter", tableText);
            Assert.Contains("Revenue", tableText);
            Assert.Contains("NetProfit", tableText);
            Assert.Contains("$1,200,000", tableText);
        }
    }

    [Fact]
    public async Task ConvertToDocxAsync_MixedPdf_PreservesEmbeddedImage()
    {
        var pdfPath = CreateMixedTextAndImagePdf();
        var shaBefore = ComputeSha256(pdfPath);

        var result = await _engine.ConvertToDocxAsync(pdfPath, "all", null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(pdfPath), shaBefore);
        Assert.NotNull(result.QualityReport);
        Assert.True(result.QualityReport.ImagesPreservedCount >= 1);

        using (var docx = WordprocessingDocument.Open(result.OutputPath, false))
        {
            Assert.NotNull(docx.MainDocumentPart?.Document?.Body);
            var drawings = docx.MainDocumentPart.Document.Body.Descendants<Drawing>().ToList();
            Assert.NotEmpty(drawings);
            Assert.Contains("Heading with Illustration", docx.MainDocumentPart.Document.Body.InnerText);
        }
    }

    [Fact]
    public async Task ConvertToDocxAsync_ScannedEmptyPdf_HonestLimitationReporting()
    {
        var pdfPath = CreateEmptyPagePdf();
        var shaBefore = ComputeSha256(pdfPath);

        var result = await _engine.ConvertToDocxAsync(pdfPath, "1", null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(pdfPath), shaBefore);
        Assert.True(result.IsScannedOrImageOnly);
        Assert.NotNull(result.QualityReport);
        Assert.Equal(PdfConversionStatus.SuccessfulWithLimitations, result.QualityReport.Status);
        Assert.NotEmpty(result.QualityReport.Warnings);
        Assert.Contains(result.QualityReport.Warnings, w => w.Contains("OCR", StringComparison.OrdinalIgnoreCase));
    }

    // ==========================================
    // 2. PDF-to-PPTX CONVERSION TESTS
    // ==========================================

    [Fact]
    public async Task ConvertToPptxAsync_StrategyA_EditableReconstruction_ValidatesPackage()
    {
        var pdfPath = CreateTwoColumnPdf();
        var shaBefore = ComputeSha256(pdfPath);

        var result = await _engine.ConvertToPptxAsync(pdfPath, PptxSlideMode.Editable, "all", null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(pdfPath), shaBefore);
        Assert.Equal(PptxSlideMode.Editable, result.SlideMode);
        Assert.NotNull(result.QualityReport);
        Assert.Equal("PptxEditableReconstruction", result.QualityReport.ConversionMode);

        var validator = new OpenXmlValidator();
        using (var pptx = PresentationDocument.Open(result.OutputPath, false))
        {
            var validationErrors = validator.Validate(pptx);
            Assert.DoesNotContain(validationErrors, e => e.ErrorType == ValidationErrorType.Schema && !e.Description.Contains("attribute"));

            Assert.NotNull(pptx.PresentationPart?.Presentation?.SlideIdList);
            var slideList = pptx.PresentationPart.Presentation.SlideIdList.ChildElements;
            Assert.Single(slideList);

            var firstSlideRelId = ((P.SlideId)slideList[0]).RelationshipId?.Value;
            Assert.NotNull(firstSlideRelId);
            var firstSlide = (SlidePart)pptx.PresentationPart.GetPartById(firstSlideRelId);
            Assert.NotNull(firstSlide.Slide);
            var textRuns = firstSlide.Slide.Descendants<A.Text>().Select(t => t.Text).ToList();
            Assert.Contains(textRuns, t => t.Contains("Executive Annual Summary"));
        }
    }

    [Fact]
    public async Task ConvertToPptxAsync_StrategyB_PageAsImage_VisualConversion_ValidatesPackage()
    {
        var pdfPath = CreateTwoColumnPdf();
        var shaBefore = ComputeSha256(pdfPath);

        var result = await _engine.ConvertToPptxAsync(pdfPath, PptxSlideMode.PageAsImage, "all", null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(pdfPath), shaBefore);
        Assert.Equal(PptxSlideMode.PageAsImage, result.SlideMode);
        Assert.NotNull(result.QualityReport);
        Assert.Equal("PptxPageAsImage", result.QualityReport.ConversionMode);
        Assert.Equal(PdfConversionStatus.Successful, result.QualityReport.Status);
        Assert.Contains("non-editable", result.SummaryNote, StringComparison.OrdinalIgnoreCase);

        var validator = new OpenXmlValidator();
        using (var pptx = PresentationDocument.Open(result.OutputPath, false))
        {
            var validationErrors = validator.Validate(pptx);
            Assert.DoesNotContain(validationErrors, e => e.ErrorType == ValidationErrorType.Schema && !e.Description.Contains("attribute"));

            Assert.NotNull(pptx.PresentationPart?.Presentation?.SlideIdList);
            var slideList = pptx.PresentationPart.Presentation.SlideIdList.ChildElements;
            Assert.Single(slideList);

            var firstSlideRelId = ((P.SlideId)slideList[0]).RelationshipId?.Value;
            Assert.NotNull(firstSlideRelId);
            var firstSlide = (SlidePart)pptx.PresentationPart.GetPartById(firstSlideRelId);
            Assert.NotNull(firstSlide.Slide);
            var pictures = firstSlide.Slide.Descendants<P.Picture>().ToList();
            Assert.NotEmpty(pictures);
        }
    }

    // ==========================================
    // 3. PDF-to-IMAGE EXPORT TESTS (PNG, JPEG, WebP)
    // ==========================================

    [Theory]
    [InlineData(ExportImageFormat.Webp, ".webp", 72)]
    [InlineData(ExportImageFormat.Webp, ".webp", 150)]
    [InlineData(ExportImageFormat.Png, ".png", 100)]
    [InlineData(ExportImageFormat.Jpeg, ".jpg", 300)]
    public async Task ExportPagesAsync_AllFormatsAndDpis_GeneratesValidImages(ExportImageFormat format, string expectedExt, int dpi)
    {
        var pdfPath = CreateTwoColumnPdf();
        var shaBefore = ComputeSha256(pdfPath);

        var result = await _engine.ExportPagesAsync(pdfPath, format, dpi, quality: 90, "1", null, CancellationToken.None);

        Assert.Single(result.ExportedFiles);
        var exportedFile = result.ExportedFiles[0];
        Assert.True(File.Exists(exportedFile));
        Assert.EndsWith(expectedExt, exportedFile, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(ComputeSha256(pdfPath), shaBefore); // Immutability

        // Verify valid image header using SKCodec
        using var codec = SKCodec.Create(exportedFile);
        Assert.NotNull(codec);
        Assert.True(codec.Info.Width > 0);
        Assert.True(codec.Info.Height > 0);

        var expectedSkFormat = format switch
        {
            ExportImageFormat.Webp => SKEncodedImageFormat.Webp,
            ExportImageFormat.Jpeg => SKEncodedImageFormat.Jpeg,
            _ => SKEncodedImageFormat.Png
        };
        Assert.Equal(expectedSkFormat, codec.EncodedFormat);
    }

    [Fact]
    public async Task ExportPagesAsync_Cancellation_CleansUpPartialsAndPromotedFiles()
    {
        var pdfPath = CreateTwoColumnPdf();
        using var cts = new CancellationTokenSource();
        cts.Cancel(); // Pre-cancel

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            _engine.ExportPagesAsync(pdfPath, ExportImageFormat.Png, 150, 85, "all", null, cts.Token));

        var partials = Directory.GetFiles(_tempDir, "*.partial", SearchOption.AllDirectories);
        Assert.Empty(partials);
    }

    // ==========================================
    // 4. PDF TEXT EXTRACTION TESTS
    // ==========================================

    [Fact]
    public async Task ExtractTextAsync_TwoColumnPdf_ExtractsInReadingOrder()
    {
        var pdfPath = CreateTwoColumnPdf();
        var shaBefore = ComputeSha256(pdfPath);

        var result = await _engine.ExtractTextAsync(pdfPath, "all", null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal(ComputeSha256(pdfPath), shaBefore);
        Assert.NotNull(result.QualityReport);
        Assert.Equal(PdfConversionStatus.Successful, result.QualityReport.Status);

        var text = File.ReadAllText(result.OutputPath, Encoding.UTF8);
        Assert.Contains("--- Page 1 ---", text);
        Assert.Contains("Executive Annual Summary", text);
        Assert.Contains("Left column paragraph line 1", text);
        Assert.Contains("Right column paragraph line 1", text);

        var leftIdx = text.IndexOf("Left column paragraph line 1", StringComparison.Ordinal);
        var rightIdx = text.IndexOf("Right column paragraph line 1", StringComparison.Ordinal);
        Assert.True(leftIdx < rightIdx);
    }

    [Fact]
    public async Task ExtractTextAsync_EmptyPdf_ReportsScannedLimitations()
    {
        var pdfPath = CreateEmptyPagePdf();

        var result = await _engine.ExtractTextAsync(pdfPath, "1", null, CancellationToken.None);

        Assert.True(File.Exists(result.OutputPath));
        Assert.True(result.IsScannedOrImageOnly);
        Assert.Equal(0, result.TotalCharactersExtracted);
        Assert.NotNull(result.QualityReport);
        Assert.Equal(PdfConversionStatus.SuccessfulWithLimitations, result.QualityReport.Status);
        Assert.NotEmpty(result.QualityReport.Warnings);
    }

    // ==========================================
    // 5. UNIFIED CONVERT & HISTORY INTEGRATION
    // ==========================================

    [Fact]
    public async Task UnifiedConvert_PdfToWebp_ExecutesSuccessfullyAndPopulatesDetails()
    {
        var pdfPath = CreateTwoColumnPdf();
        var classified = _classifier.Classify(pdfPath);
        var catalog = new ProcessorCatalog(_classifier, _allocator);
        var capabilities = new SmartShift.Core.Capabilities.FileCapabilityRegistry();
        var conversionService = new SmartShift.Processing.Conversion.UnifiedConversionService(capabilities, catalog, _allocator, _classifier);

        var options = new Dictionary<string, string>
        {
            ["format"] = "webp",
            ["dpi"] = "150",
            ["quality"] = "85",
            ["pageRange"] = "1"
        };

        var result = await conversionService.ConvertAsync(classified, DetectedFormat.Webp, options, _tempDir, null, CancellationToken.None);

        Assert.True(result.Succeeded);
        Assert.True(result.OriginalUnchanged);
        Assert.NotNull(result.OutputPath);
        Assert.True(File.Exists(result.OutputPath));
        Assert.EndsWith(".webp", result.OutputPath, StringComparison.OrdinalIgnoreCase);

        // Verify multi-output paths detail exists
        Assert.True(result.Details.ContainsKey("outputPaths"));
        Assert.Equal("WebP", result.Details["outputFormat"]);
    }
}
