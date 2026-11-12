using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf;
using SkiaSharp;
using SmartShift.Core.Actions;
using SmartShift.Core.Classification;
using SmartShift.Processing.Quality;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace SmartShift.Processing.Tests;

public sealed class ConversionQualityAndFidelityTests : IDisposable
{
    private readonly string _tempDir;
    private readonly StructuralFidelityValidator _structuralValidator;
    private readonly ContentIntegrityValidator _contentValidator;
    private readonly VisualFidelityAnalyzer _visualAnalyzer;
    private readonly ConversionQualityAnalyzer _qualityAnalyzer;

    public ConversionQualityAndFidelityTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "smartshift-phase3-step4-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
        _structuralValidator = new StructuralFidelityValidator();
        _contentValidator = new ContentIntegrityValidator();
        _visualAnalyzer = new VisualFidelityAnalyzer();
        _qualityAnalyzer = new ConversionQualityAnalyzer(_structuralValidator, _contentValidator, _visualAnalyzer);
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

    private string CreateDocx(string name, string text, bool addBrokenImage = false)
    {
        var path = Path.Combine(_tempDir, name);
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new W.Document();
            var body = mainPart.Document.AppendChild(new W.Body());

            body.AppendChild(new W.Paragraph(new W.Run(new W.Text(text))));

            if (addBrokenImage)
            {
                var drawing = new W.Drawing(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 990000L, Cy = 792000L },
                        new DocumentFormat.OpenXml.Drawing.Graphic(
                            new DocumentFormat.OpenXml.Drawing.GraphicData(
                                new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                        new A.Blip { Embed = "rIdNonExistent" }
                                    )
                                )
                            ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                        )
                    )
                );
                body.AppendChild(new W.Paragraph(new W.Run(drawing)));
            }

            mainPart.Document.Save();
        }
        return path;
    }

    private string CreatePptx(string name, string slideText)
    {
        var path = Path.Combine(_tempDir, name);
        using (var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation))
        {
            var presPart = doc.AddPresentationPart();
            presPart.Presentation = new Presentation();

            var slidePart = presPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 1, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(),
                        new P.Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties { Id = 2, Name = "Text 1" },
                                new NonVisualShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new ShapeProperties(),
                            new TextBody(
                                new A.BodyProperties(),
                                new A.ListStyle(),
                                new A.Paragraph(new A.Run(new A.Text(slideText)))
                            )
                        )
                    )
                )
            );
            slidePart.Slide.Save();

            presPart.Presentation.SlideIdList = new SlideIdList(new SlideId { Id = 256, RelationshipId = presPart.GetIdOfPart(slidePart) });
            presPart.Presentation.Save();
        }
        return path;
    }

    private string CreatePdf(string name, int pageCount = 1)
    {
        var path = Path.Combine(_tempDir, name);
        using (var doc = new PdfDocument())
        {
            for (int i = 0; i < pageCount; i++)
            {
                var page = doc.AddPage();
                using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page);
                var font = new PdfSharp.Drawing.XFont("Arial", 12, PdfSharp.Drawing.XFontStyleEx.Regular);
                gfx.DrawString($"Page {i + 1} content", font, PdfSharp.Drawing.XBrushes.Black, 50, 50);
            }
            doc.Save(path);
        }
        return path;
    }

    private string CreatePng(string name, int width, int height, SKColor color)
    {
        var path = Path.Combine(_tempDir, name);
        using var bitmap = new SKBitmap(width, height);
        using (var canvas = new SKCanvas(bitmap))
        {
            canvas.Clear(color);
        }
        using var img = SKImage.FromBitmap(bitmap);
        using var data = img.Encode(SKEncodedImageFormat.Png, 100);
        using var fs = File.OpenWrite(path);
        data.SaveTo(fs);
        return path;
    }

    // ==========================================
    // STRUCTURAL VALIDATION TESTS
    // ==========================================

    [Fact]
    public void StructuralValidator_ValidDocx_PassesValidation()
    {
        var docxPath = CreateDocx("valid.docx", "Valid DOCX document paragraph.");
        var metrics = _structuralValidator.Validate(docxPath, DetectedFormat.Docx);

        Assert.True(metrics.PackageValid);
        Assert.True(metrics.ParagraphCount >= 1);
        Assert.Empty(metrics.StructuralWarnings);
    }

    [Fact]
    public void StructuralValidator_DocxWithBrokenImage_DetectsBrokenRelationship()
    {
        var docxPath = CreateDocx("broken_img.docx", "Document with broken image.", addBrokenImage: true);
        var metrics = _structuralValidator.Validate(docxPath, DetectedFormat.Docx);

        Assert.False(metrics.PackageValid);
        Assert.Contains(metrics.StructuralWarnings, w => w.Contains("Broken image relationship detected"));
    }

    [Fact]
    public void StructuralValidator_ValidPptx_PassesValidation()
    {
        var pptxPath = CreatePptx("valid.pptx", "Slide presentation text.");
        var metrics = _structuralValidator.Validate(pptxPath, DetectedFormat.Pptx);

        Assert.True(metrics.PackageValid);
        Assert.Equal(1, metrics.PageOrSlideCount);
        Assert.True(metrics.ShapeCount >= 1);
        Assert.Empty(metrics.StructuralWarnings);
    }

    [Fact]
    public void StructuralValidator_ValidPdf_PassesValidation()
    {
        var pdfPath = CreatePdf("valid.pdf", pageCount: 3);
        var metrics = _structuralValidator.Validate(pdfPath, DetectedFormat.Pdf);

        Assert.True(metrics.PackageValid);
        Assert.Equal(3, metrics.PageOrSlideCount);
        Assert.Empty(metrics.StructuralWarnings);
    }

    [Fact]
    public void StructuralValidator_ValidImage_PassesValidation()
    {
        var imgPath = CreatePng("test.png", 200, 150, SKColors.Blue);
        var metrics = _structuralValidator.Validate(imgPath, DetectedFormat.Png);

        Assert.True(metrics.PackageValid);
        Assert.Equal(1, metrics.ImageCount);
        Assert.Empty(metrics.StructuralWarnings);
    }

    [Fact]
    public void StructuralValidator_NonExistentFile_ReturnsInvalid()
    {
        var metrics = _structuralValidator.Validate(Path.Combine(_tempDir, "missing.docx"), DetectedFormat.Docx);
        Assert.False(metrics.PackageValid);
        Assert.NotEmpty(metrics.StructuralWarnings);
    }

    // ==========================================
    // CONTENT INTEGRITY TESTS
    // ==========================================

    [Fact]
    public void ContentValidator_IdenticalText_ReturnsSimilarity1()
    {
        var srcPath = Path.Combine(_tempDir, "src.txt");
        var outPath = Path.Combine(_tempDir, "out.txt");
        File.WriteAllText(srcPath, "The quick brown fox jumps over the lazy dog.");
        File.WriteAllText(outPath, "The quick brown fox jumps over the lazy dog.");

        var metrics = _contentValidator.Validate(srcPath, DetectedFormat.Txt, outPath, DetectedFormat.Txt);

        Assert.NotNull(metrics);
        Assert.Equal(1.0, metrics!.TextSimilarityScore);
        Assert.Equal(0, metrics.MissingTextSegments);
        Assert.True(metrics.UnicodePreserved);
    }

    [Fact]
    public void ContentValidator_MinorFormattingDifference_ReturnsHighSimilarity()
    {
        var srcPath = Path.Combine(_tempDir, "src_ws.txt");
        var outPath = Path.Combine(_tempDir, "out_ws.txt");
        File.WriteAllText(srcPath, "Section 1:\n\nThe quick brown fox   jumps over the lazy dog.");
        File.WriteAllText(outPath, "Section 1: The quick brown fox jumps over the lazy dog.");

        var metrics = _contentValidator.Validate(srcPath, DetectedFormat.Txt, outPath, DetectedFormat.Txt);

        Assert.NotNull(metrics);
        Assert.True(metrics!.TextSimilarityScore >= 0.95);
        Assert.Equal(0, metrics.MissingTextSegments);
    }

    [Fact]
    public void ContentValidator_SignificantContentLoss_DetectsMissingWords()
    {
        var srcPath = Path.Combine(_tempDir, "src_full.txt");
        var outPath = Path.Combine(_tempDir, "out_partial.txt");
        File.WriteAllText(srcPath, "SmartShift enterprise document conversion engine with verified quality and fidelity benchmarks.");
        File.WriteAllText(outPath, "SmartShift enterprise.");

        var metrics = _contentValidator.Validate(srcPath, DetectedFormat.Txt, outPath, DetectedFormat.Txt);

        Assert.NotNull(metrics);
        Assert.True(metrics!.TextSimilarityScore < 0.50);
        Assert.True(metrics.MissingTextSegments > 0);
        Assert.Contains(metrics.ContentWarnings, w => w.Contains("missing in output"));
    }

    [Fact]
    public void ContentValidator_UnicodePreservation_DetectsMissingCharacters()
    {
        var srcPath = Path.Combine(_tempDir, "unicode_src.txt");
        var outPath = Path.Combine(_tempDir, "unicode_out.txt");
        File.WriteAllText(srcPath, "International review: München café façade — 100€");
        File.WriteAllText(outPath, "International review: Munchen cafe facade - 100");

        var metrics = _contentValidator.Validate(srcPath, DetectedFormat.Txt, outPath, DetectedFormat.Txt);

        Assert.NotNull(metrics);
        Assert.False(metrics!.UnicodePreserved);
        Assert.Contains(metrics.ContentWarnings, w => w.Contains("non-ASCII Unicode characters"));
    }

    // ==========================================
    // VISUAL FIDELITY TESTS
    // ==========================================

    [Fact]
    public void VisualAnalyzer_IdenticalImages_ReturnsZeroError()
    {
        using var bm1 = new SKBitmap(100, 100);
        using var bm2 = new SKBitmap(100, 100);
        using (var canvas1 = new SKCanvas(bm1)) canvas1.Clear(SKColors.Red);
        using (var canvas2 = new SKCanvas(bm2)) canvas2.Clear(SKColors.Red);

        var metrics = VisualFidelityAnalyzer.CompareBitmaps(bm1, bm2);

        Assert.Equal(VisualComparisonStatus.Comparable, metrics.Status);
        Assert.Equal(0.0, metrics.MeanAbsoluteError);
        Assert.Equal(0.0, metrics.RootMeanSquareError);
        Assert.Equal(0.0, metrics.PixelDifferencePercentage);
    }

    [Fact]
    public void VisualAnalyzer_DifferingDimensions_ResamplesAndCompares()
    {
        using var bm1 = new SKBitmap(100, 100);
        using var bm2 = new SKBitmap(50, 50);
        using (var canvas1 = new SKCanvas(bm1)) canvas1.Clear(SKColors.Green);
        using (var canvas2 = new SKCanvas(bm2)) canvas2.Clear(SKColors.Green);

        var metrics = VisualFidelityAnalyzer.CompareBitmaps(bm1, bm2);

        Assert.Equal(VisualComparisonStatus.DifferingDimensions, metrics.Status);
        Assert.True(metrics.MeanAbsoluteError < 0.05);
        Assert.Contains(metrics.VisualWarnings, w => w.Contains("Dimension discrepancy"));
    }

    [Fact]
    public void VisualAnalyzer_CompletelyDifferentColors_DetectsPixelDiff()
    {
        using var bm1 = new SKBitmap(100, 100);
        using var bm2 = new SKBitmap(100, 100);
        using (var canvas1 = new SKCanvas(bm1)) canvas1.Clear(SKColors.White);
        using (var canvas2 = new SKCanvas(bm2)) canvas2.Clear(SKColors.Black);

        var metrics = VisualFidelityAnalyzer.CompareBitmaps(bm1, bm2);

        Assert.Equal(100.0, metrics.PixelDifferencePercentage);
        Assert.True(metrics.MeanAbsoluteError >= 0.70);
    }

    [Fact]
    public void VisualAnalyzer_NonRenderableFormat_ReportsNotComparable()
    {
        var docxPath = CreateDocx("test.docx", "Some text");
        var metrics = _visualAnalyzer.Analyze(docxPath, DetectedFormat.Docx, docxPath, DetectedFormat.Docx);

        Assert.Equal(VisualComparisonStatus.NotComparable, metrics.Status);
        Assert.NotNull(metrics.Rationale);
        Assert.Contains("not supported", metrics.Rationale);
    }

    // ==========================================
    // CAPABILITY MATRIX & CLASSIFICATION TESTS
    // ==========================================

    [Fact]
    public void ConversionCapabilityMatrix_ContainsCoreWorkflows()
    {
        var all = ConversionCapabilityMatrix.GetAllDescriptors();
        Assert.True(all.Count >= 10);

        var pdfDocx = ConversionCapabilityMatrix.GetDescriptor(ActionIds.ConvertPdfToDocx);
        Assert.NotNull(pdfDocx);
        Assert.True(pdfDocx!.SupportsStructuralValidation);
        Assert.True(pdfDocx.SupportsContentValidation);
        Assert.False(pdfDocx.SupportsVisualValidation);
        Assert.NotEmpty(pdfDocx.KnownLimitations);

        var pptxPdf = ConversionCapabilityMatrix.GetDescriptor(ActionIds.ConvertPptxToPdf);
        Assert.NotNull(pptxPdf);
        Assert.True(pptxPdf!.SupportsVisualValidation);
    }

    [Fact]
    public async Task QualityAnalyzer_ValidConversion_ReturnsValidated()
    {
        var srcTxt = Path.Combine(_tempDir, "qa_src.txt");
        var outTxt = Path.Combine(_tempDir, "qa_out.txt");
        File.WriteAllText(srcTxt, "Enterprise system conversion test.");
        File.WriteAllText(outTxt, "Enterprise system conversion test.");

        var report = await _qualityAnalyzer.AnalyzeAsync(
            srcTxt,
            DetectedFormat.Txt,
            outTxt,
            DetectedFormat.Txt,
            ActionIds.UnifiedConvert,
            ValidationLevel.Standard);

        Assert.Equal(QualityClassification.Validated, report.Classification);
        Assert.True(report.CompositeConfidenceScore >= 0.95);
        Assert.True(report.IsAcceptable);
        Assert.Empty(report.Warnings);
    }

    [Fact]
    public async Task QualityAnalyzer_KnownLimitations_ReturnsValidatedWithLimitations()
    {
        var srcPdf = CreatePdf("lim_src.pdf", pageCount: 1);
        var outDocx = CreateDocx("lim_out.docx", "Converted paragraph content from PDF.");

        var report = await _qualityAnalyzer.AnalyzeAsync(
            srcPdf,
            DetectedFormat.Pdf,
            outDocx,
            DetectedFormat.Docx,
            ActionIds.ConvertPdfToDocx,
            ValidationLevel.Standard);

        Assert.Equal(QualityClassification.ValidatedWithLimitations, report.Classification);
        Assert.True(report.IsAcceptable);
        Assert.NotEmpty(report.Limitations);
    }

    [Fact]
    public async Task QualityAnalyzer_StructuralMismatch_ReturnsStructuralMismatch()
    {
        var srcTxt = Path.Combine(_tempDir, "valid_src.txt");
        var brokenDocx = Path.Combine(_tempDir, "corrupt.docx");
        File.WriteAllText(srcTxt, "Hello world");
        File.WriteAllText(brokenDocx, "Not an OpenXML zip package content");

        var report = await _qualityAnalyzer.AnalyzeAsync(
            srcTxt,
            DetectedFormat.Txt,
            brokenDocx,
            DetectedFormat.Docx,
            ActionIds.ConvertPdfToDocx,
            ValidationLevel.Basic);

        Assert.Equal(QualityClassification.StructuralMismatch, report.Classification);
        Assert.False(report.IsAcceptable);
        Assert.Equal(0.0, report.CompositeConfidenceScore);
    }
}
