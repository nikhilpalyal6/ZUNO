using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using SmartShift.Core.Actions;
using SmartShift.Core.Capabilities;
using SmartShift.Core.Classification;
using SmartShift.Core.IO;
using SmartShift.Infrastructure.Output;
using SmartShift.Processing;
using SmartShift.Processing.Documents.Editing;
using SmartShift.Processing.Pdf.Editing;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace SmartShift.Processing.Tests;

public sealed class DocumentEditingAndTransformationTests : IDisposable
{
    private readonly string _tempDir;
    private readonly FileClassifier _classifier;
    private readonly OutputPathAllocator _allocator;
    private readonly DocxMergeService _docxMergeService;
    private readonly DocxPageNumberService _docxPageNumberService;
    private readonly DocxWatermarkService _docxWatermarkService;
    private readonly DocxMetadataService _docxMetadataService;
    private readonly PdfWatermarkService _pdfWatermarkService;
    private readonly PdfMetadataService _pdfMetadataService;
    private readonly ProcessorCatalog _catalog;

    public DocumentEditingAndTransformationTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "smartshift-phase3-step3-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
        _classifier = new FileClassifier();
        _allocator = new OutputPathAllocator(_tempDir);
        _docxMergeService = new DocxMergeService(_allocator);
        _docxPageNumberService = new DocxPageNumberService(_allocator);
        _docxWatermarkService = new DocxWatermarkService(_allocator);
        _docxMetadataService = new DocxMetadataService(_allocator);
        _pdfWatermarkService = new PdfWatermarkService(_allocator);
        _pdfMetadataService = new PdfMetadataService(_allocator);
        _catalog = new ProcessorCatalog(_classifier, _allocator);
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

    private string CreateSampleDocx(string name, int paragraphCount = 2, bool includeTable = false, bool includeImage = false)
    {
        var path = Path.Combine(_tempDir, name);
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new W.Document();
            var body = mainPart.Document.AppendChild(new W.Body());

            for (int i = 1; i <= paragraphCount; i++)
            {
                body.AppendChild(new W.Paragraph(new W.Run(new W.Text($"Document '{name}' paragraph {i}."))));
            }

            if (includeTable)
            {
                var table = new W.Table();
                var row = new W.TableRow();
                row.AppendChild(new W.TableCell(new W.Paragraph(new W.Run(new W.Text("Cell 1")))));
                row.AppendChild(new W.TableCell(new W.Paragraph(new W.Run(new W.Text("Cell 2")))));
                table.AppendChild(row);
                body.AppendChild(table);
            }

            if (includeImage)
            {
                var imgPart = mainPart.AddImagePart(ImagePartType.Png);
                var fakeBytes = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
                using (var ms = new MemoryStream(fakeBytes))
                {
                    imgPart.FeedData(ms);
                }
                var relId = mainPart.GetIdOfPart(imgPart);

                var drawing = new W.Drawing(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 990000L, Cy = 792000L },
                        new DocumentFormat.OpenXml.Drawing.Graphic(
                            new DocumentFormat.OpenXml.Drawing.GraphicData(
                                new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                        new A.Blip { Embed = relId }
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

    private string CreateSamplePdf(string name, int pageCount = 2)
    {
        var path = Path.Combine(_tempDir, name);
        using (var doc = new PdfDocument())
        {
            for (int i = 1; i <= pageCount; i++)
            {
                var page = doc.AddPage();
                using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page);
                var font = new PdfSharp.Drawing.XFont("Arial", 14, PdfSharp.Drawing.XFontStyleEx.Regular);
                gfx.DrawString($"Page {i} content", font, PdfSharp.Drawing.XBrushes.Black, 50, 50);
            }
            doc.Save(path);
        }
        return path;
    }

    // ==========================================
    // DOCX MERGE TESTS
    // ==========================================

    [Fact]
    public async Task DocxMerge_MergesMultipleFiles_PreservesContentAndDoesNotMutateSources()
    {
        var doc1 = CreateSampleDocx("doc1.docx", paragraphCount: 3, includeTable: true, includeImage: false);
        var doc2 = CreateSampleDocx("doc2.docx", paragraphCount: 2, includeTable: false, includeImage: true);

        var preHash1 = ComputeSha256(doc1);
        var preHash2 = ComputeSha256(doc2);
        var outputPath = Path.Combine(_tempDir, "merged_output.docx");

        var result = await _docxMergeService.MergeAsync([doc1, doc2], outputPath);

        Assert.True(result.IsSuccess);
        Assert.True(File.Exists(outputPath));
        Assert.Equal(2, result.DocumentCount);
        Assert.True(result.ParagraphCount >= 5);
        Assert.True(result.TableCount >= 1);
        Assert.True(result.ImageCount >= 1);

        // Verify output document integrity
        using (var targetDoc = WordprocessingDocument.Open(outputPath, false))
        {
            var main = targetDoc.MainDocumentPart;
            Assert.NotNull(main);
            Assert.NotNull(main!.Document);
            var body = main.Document.Body;
            Assert.NotNull(body);
            var paragraphs = body!.Elements<W.Paragraph>().ToList();
            Assert.True(paragraphs.Count >= 5);
            var tables = body.Elements<W.Table>().ToList();
            Assert.True(tables.Count >= 1);
            Assert.True(main.ImageParts.Any());
        }

        // Verify non-destructive source integrity
        Assert.Equal(preHash1, ComputeSha256(doc1));
        Assert.Equal(preHash2, ComputeSha256(doc2));
    }

    [Fact]
    public async Task DocxMerge_LessThanTwoFiles_ThrowsArgumentException()
    {
        var doc1 = CreateSampleDocx("single.docx");
        var outputPath = Path.Combine(_tempDir, "invalid_merge.docx");

        await Assert.ThrowsAsync<ArgumentException>(() => _docxMergeService.MergeAsync([doc1], outputPath));
    }

    // ==========================================
    // DOCX PAGE NUMBERING TESTS
    // ==========================================

    [Fact]
    public async Task DocxPageNumber_FooterCenter_PageNumberTotal_AddsFields()
    {
        var docPath = CreateSampleDocx("page_num_test.docx", paragraphCount: 3);
        var preHash = ComputeSha256(docPath);
        var outputPath = Path.Combine(_tempDir, "page_num_output.docx");

        var options = new DocxPageNumberOptions(
            Position: PageNumberPosition.Footer,
            Alignment: PageNumberAlignment.Center,
            Format: PageNumberFormat.PageNumberTotal);

        var result = await _docxPageNumberService.AddPageNumbersAsync(docPath, outputPath, options);

        Assert.True(result.IsSuccess);
        Assert.True(File.Exists(outputPath));

        using (var doc = WordprocessingDocument.Open(outputPath, false))
        {
            var mainPart = doc.MainDocumentPart!;
            Assert.True(mainPart.FooterParts.Any());

            var footer = mainPart.FooterParts.First().Footer;
            Assert.NotNull(footer);
            var text = footer.InnerText;
            Assert.Contains("Page", text);
            var fields = footer.Descendants<W.SimpleField>().ToList();
            Assert.Equal(2, fields.Count);
            Assert.Contains(fields, f => f.Instruction?.Value == "PAGE");
            Assert.Contains(fields, f => f.Instruction?.Value == "NUMPAGES");
        }

        Assert.Equal(preHash, ComputeSha256(docPath));
    }

    [Fact]
    public async Task DocxPageNumber_HeaderLeft_NumberOnly_AddsField()
    {
        var docPath = CreateSampleDocx("page_num_header_test.docx", paragraphCount: 2);
        var outputPath = Path.Combine(_tempDir, "page_num_header_output.docx");

        var options = new DocxPageNumberOptions(
            Position: PageNumberPosition.Header,
            Alignment: PageNumberAlignment.Left,
            Format: PageNumberFormat.NumberOnly);

        var result = await _docxPageNumberService.AddPageNumbersAsync(docPath, outputPath, options);

        Assert.True(result.IsSuccess);
        Assert.True(File.Exists(outputPath));

        using (var doc = WordprocessingDocument.Open(outputPath, false))
        {
            var mainPart = doc.MainDocumentPart!;
            Assert.True(mainPart.HeaderParts.Any());
            var header = mainPart.HeaderParts.First().Header;
            Assert.NotNull(header);
            var field = header.Descendants<W.SimpleField>().Single();
            Assert.Equal("PAGE", field.Instruction?.Value);
        }
    }

    // ==========================================
    // DOCX WATERMARK TESTS
    // ==========================================

    [Fact]
    public async Task DocxWatermark_EmbedsWatermarkInHeader_DoesNotMutateSource()
    {
        var docPath = CreateSampleDocx("watermark_test.docx", paragraphCount: 2);
        var preHash = ComputeSha256(docPath);
        var outputPath = Path.Combine(_tempDir, "watermark_output.docx");

        var options = new DocxWatermarkOptions("CONFIDENTIAL", FontSize: 54, Opacity: 0.3, IsDiagonal: true);
        var result = await _docxWatermarkService.AddWatermarkAsync(docPath, outputPath, options);

        Assert.True(result.IsSuccess);
        Assert.True(File.Exists(outputPath));

        using (var doc = WordprocessingDocument.Open(outputPath, false))
        {
            var mainPart = doc.MainDocumentPart!;
            Assert.True(mainPart.HeaderParts.Any());
            var header = mainPart.HeaderParts.First().Header;
            Assert.NotNull(header);
            var shape = header.Descendants<DocumentFormat.OpenXml.Vml.Shape>().FirstOrDefault();
            Assert.NotNull(shape);
            Assert.Equal("WordWatermark", shape.Id?.Value);
            var textPath = shape.Descendants<DocumentFormat.OpenXml.Vml.TextPath>().FirstOrDefault();
            Assert.NotNull(textPath);
            Assert.Equal("CONFIDENTIAL", textPath.String?.Value);
        }

        Assert.Equal(preHash, ComputeSha256(docPath));
    }

    [Fact]
    public async Task DocxWatermark_EmptyText_ThrowsArgumentException()
    {
        var docPath = CreateSampleDocx("empty_wm.docx");
        var outputPath = Path.Combine(_tempDir, "wm_err.docx");

        await Assert.ThrowsAsync<ArgumentException>(() =>
            _docxWatermarkService.AddWatermarkAsync(docPath, outputPath, new DocxWatermarkOptions("")));
    }

    // ==========================================
    // DOCX METADATA TESTS
    // ==========================================

    [Fact]
    public async Task DocxMetadata_ReadAndUpdate_PreservesFileAndUpdatesProperties()
    {
        var docPath = CreateSampleDocx("meta_test.docx", paragraphCount: 1);
        var preHash = ComputeSha256(docPath);
        var outputPath = Path.Combine(_tempDir, "meta_output.docx");

        var metadata = new DocxMetadata(
            Title: "Quarterly Audit",
            Subject: "Finance Review",
            Author: "Jane Doe",
            Keywords: "audit; finance; report",
            Comments: "Final approved draft",
            Category: "Finance");

        var updateResult = await _docxMetadataService.UpdateMetadataAsync(docPath, outputPath, metadata);
        Assert.True(updateResult.IsSuccess);
        Assert.True(File.Exists(outputPath));

        var readBack = await _docxMetadataService.ReadMetadataAsync(outputPath);
        Assert.Equal("Quarterly Audit", readBack.Title);
        Assert.Equal("Finance Review", readBack.Subject);
        Assert.Equal("Jane Doe", readBack.Author);
        Assert.Equal("audit; finance; report", readBack.Keywords);
        Assert.Equal("Final approved draft", readBack.Comments);
        Assert.Equal("Finance", readBack.Category);

        Assert.Equal(preHash, ComputeSha256(docPath));
    }

    // ==========================================
    // PDF WATERMARK TESTS
    // ==========================================

    [Fact]
    public async Task PdfWatermark_AppliesWatermarkToAllPages_PreservesSource()
    {
        var pdfPath = CreateSamplePdf("pdf_wm_test.pdf", pageCount: 3);
        var preHash = ComputeSha256(pdfPath);
        var outputPath = Path.Combine(_tempDir, "pdf_wm_output.pdf");

        var options = new PdfWatermarkOptions(
            Text: "RESTRICTED",
            FontSize: 40,
            Opacity: 0.35,
            IsDiagonal: true,
            HexColor: "#FF0000",
            PageRange: "all");

        var result = await _pdfWatermarkService.AddWatermarkAsync(pdfPath, outputPath, options);

        Assert.True(result.IsSuccess);
        Assert.True(File.Exists(outputPath));

        using (var doc = PdfReader.Open(outputPath, PdfDocumentOpenMode.Import))
        {
            Assert.Equal(3, doc.PageCount);
        }

        Assert.Equal(preHash, ComputeSha256(pdfPath));
    }

    [Fact]
    public async Task PdfWatermark_AppliesWatermarkToOddPages()
    {
        var pdfPath = CreateSamplePdf("pdf_odd_wm.pdf", pageCount: 4);
        var outputPath = Path.Combine(_tempDir, "pdf_odd_wm_output.pdf");

        var options = new PdfWatermarkOptions(
            Text: "COPY",
            FontSize: 30,
            Opacity: 0.2,
            IsDiagonal: false,
            PageRange: "odd");

        var result = await _pdfWatermarkService.AddWatermarkAsync(pdfPath, outputPath, options);

        Assert.True(result.IsSuccess);
        Assert.True(File.Exists(outputPath));

        using (var doc = PdfReader.Open(outputPath, PdfDocumentOpenMode.Import))
        {
            Assert.Equal(4, doc.PageCount);
        }
    }

    // ==========================================
    // PDF METADATA TESTS
    // ==========================================

    [Fact]
    public async Task PdfMetadata_ReadAndUpdate_UpdatesInfoDictionary()
    {
        var pdfPath = CreateSamplePdf("pdf_meta_test.pdf", pageCount: 1);
        var preHash = ComputeSha256(pdfPath);
        var outputPath = Path.Combine(_tempDir, "pdf_meta_output.pdf");

        var metadata = new SmartShift.Processing.Pdf.Editing.PdfMetadata(
            Title: "Annual Security Report",
            Author: "John Smith",
            Subject: "Cybersecurity",
            Keywords: "security, report, 2026",
            Creator: "SmartShift Engine",
            Producer: "SmartShift Security");

        var updateResult = await _pdfMetadataService.UpdateMetadataAsync(pdfPath, outputPath, metadata);
        Assert.True(updateResult.IsSuccess);
        Assert.True(File.Exists(outputPath));

        var readBack = await _pdfMetadataService.ReadMetadataAsync(outputPath);
        Assert.Equal("Annual Security Report", readBack.Title);
        Assert.Equal("John Smith", readBack.Author);
        Assert.Equal("Cybersecurity", readBack.Subject);
        Assert.Equal("security, report, 2026", readBack.Keywords);
        Assert.Equal("SmartShift Engine", readBack.Creator);

        Assert.Equal(preHash, ComputeSha256(pdfPath));
    }

    // ==========================================
    // PROCESSOR INTEGRATION & PIPELINE TESTS
    // ==========================================

    [Fact]
    public async Task DocxMergeProcessor_ViaCatalog_Success()
    {
        var doc1 = CreateSampleDocx("p_doc1.docx", paragraphCount: 2);
        var doc2 = CreateSampleDocx("p_doc2.docx", paragraphCount: 2);

        var processor = _catalog.Find(ActionIds.MergeDocx);
        Assert.NotNull(processor);

        var classified = _classifier.Classify(doc1);
        var request = new ProcessingRequest(
            classified,
            ActionIds.MergeDocx,
            new ProcessingOptions(new Dictionary<string, string>
            {
                ["additionalFiles"] = doc2,
                ["outputFileName"] = "p_merged.docx"
            }),
            _tempDir);

        var result = await processor.ProcessAsync(request, null, CancellationToken.None);
        Assert.True(result.Succeeded);
        Assert.NotNull(result.OutputPath);
        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal("2", result.Details["TotalFiles"]);
    }

    [Fact]
    public async Task DocxPageNumberProcessor_ViaCatalog_Success()
    {
        var doc = CreateSampleDocx("p_num.docx", paragraphCount: 2);
        var processor = _catalog.Find(ActionIds.AddDocxPageNumbers);
        Assert.NotNull(processor);

        var classified = _classifier.Classify(doc);
        var request = new ProcessingRequest(
            classified,
            ActionIds.AddDocxPageNumbers,
            new ProcessingOptions(new Dictionary<string, string>
            {
                ["position"] = "footer",
                ["alignment"] = "center",
                ["format"] = "pageNumberTotal"
            }),
            _tempDir);

        var result = await processor.ProcessAsync(request, null, CancellationToken.None);
        Assert.True(result.Succeeded);
        Assert.NotNull(result.OutputPath);
        Assert.True(File.Exists(result.OutputPath));
    }

    [Fact]
    public async Task DocxWatermarkProcessor_ViaCatalog_Success()
    {
        var doc = CreateSampleDocx("p_wm.docx", paragraphCount: 1);
        var processor = _catalog.Find(ActionIds.AddDocxWatermark);
        Assert.NotNull(processor);

        var classified = _classifier.Classify(doc);
        var request = new ProcessingRequest(
            classified,
            ActionIds.AddDocxWatermark,
            new ProcessingOptions(new Dictionary<string, string>
            {
                ["text"] = "SAMPLE",
                ["opacity"] = "0.2"
            }),
            _tempDir);

        var result = await processor.ProcessAsync(request, null, CancellationToken.None);
        Assert.True(result.Succeeded);
        Assert.NotNull(result.OutputPath);
        Assert.True(File.Exists(result.OutputPath));
    }

    [Fact]
    public async Task DocxMetadataProcessor_ViaCatalog_Success()
    {
        var doc = CreateSampleDocx("p_meta.docx", paragraphCount: 1);
        var processor = _catalog.Find(ActionIds.EditDocxMetadata);
        Assert.NotNull(processor);

        var classified = _classifier.Classify(doc);
        var request = new ProcessingRequest(
            classified,
            ActionIds.EditDocxMetadata,
            new ProcessingOptions(new Dictionary<string, string>
            {
                ["title"] = "Catalog Test Title",
                ["author"] = "Catalog Author"
            }),
            _tempDir);

        var result = await processor.ProcessAsync(request, null, CancellationToken.None);
        Assert.True(result.Succeeded);
        Assert.NotNull(result.OutputPath);
        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal("Catalog Test Title", result.Details["Title"]);
    }

    [Fact]
    public async Task PdfWatermarkProcessor_ViaCatalog_Success()
    {
        var pdf = CreateSamplePdf("p_pdf_wm.pdf", pageCount: 2);
        var processor = _catalog.Find(ActionIds.WatermarkPdf);
        Assert.NotNull(processor);

        var classified = _classifier.Classify(pdf);
        var request = new ProcessingRequest(
            classified,
            ActionIds.WatermarkPdf,
            new ProcessingOptions(new Dictionary<string, string>
            {
                ["text"] = "OFFICIAL",
                ["opacity"] = "0.3"
            }),
            _tempDir);

        var result = await processor.ProcessAsync(request, null, CancellationToken.None);
        Assert.True(result.Succeeded);
        Assert.NotNull(result.OutputPath);
        Assert.True(File.Exists(result.OutputPath));
    }

    [Fact]
    public async Task PdfMetadataProcessor_ViaCatalog_Success()
    {
        var pdf = CreateSamplePdf("p_pdf_meta.pdf", pageCount: 1);
        var processor = _catalog.Find(ActionIds.EditPdfMetadata);
        Assert.NotNull(processor);

        var classified = _classifier.Classify(pdf);
        var request = new ProcessingRequest(
            classified,
            ActionIds.EditPdfMetadata,
            new ProcessingOptions(new Dictionary<string, string>
            {
                ["title"] = "PDF Catalog Title",
                ["author"] = "PDF Catalog Author"
            }),
            _tempDir);

        var result = await processor.ProcessAsync(request, null, CancellationToken.None);
        Assert.True(result.Succeeded);
        Assert.NotNull(result.OutputPath);
        Assert.True(File.Exists(result.OutputPath));
        Assert.Equal("PDF Catalog Title", result.Details["Title"]);
    }

    [Fact]
    public void ActionOptionValidator_ValidatesWatermarkAndPageNumbers()
    {
        // Watermark validation
        var wmErrors = ActionOptionValidator.Validate(ActionIds.WatermarkPdf, new Dictionary<string, string>
        {
            ["text"] = "  ",
            ["opacity"] = "1.5"
        });
        Assert.Contains(wmErrors, e => e.Contains("Watermark text cannot be empty"));
        Assert.Contains(wmErrors, e => e.Contains("opacity must be between 0.01 and 1.0"));

        // Page number validation
        var pnErrors = ActionOptionValidator.Validate(ActionIds.AddDocxPageNumbers, new Dictionary<string, string>
        {
            ["position"] = "invalid_pos",
            ["alignment"] = "invalid_align"
        });
        Assert.Contains(pnErrors, e => e.Contains("position must be either 'footer' or 'header'"));
        Assert.Contains(pnErrors, e => e.Contains("alignment must be 'left', 'center', or 'right'"));
    }

    [Fact]
    public void FileCapabilityRegistry_ExposesNewEditingActionsForDocxAndPdf()
    {
        var registry = new FileCapabilityRegistry();

        var docxDesc = registry.GetDescriptor(DetectedFormat.Docx);
        Assert.True(docxDesc.CanEdit);
        Assert.Contains(ActionIds.MergeDocx, docxDesc.SupportedActionIds);
        Assert.Contains(ActionIds.AddDocxPageNumbers, docxDesc.SupportedActionIds);
        Assert.Contains(ActionIds.AddDocxWatermark, docxDesc.SupportedActionIds);
        Assert.Contains(ActionIds.EditDocxMetadata, docxDesc.SupportedActionIds);

        var pdfDesc = registry.GetDescriptor(DetectedFormat.Pdf);
        Assert.True(pdfDesc.CanEdit);
        Assert.Contains(ActionIds.WatermarkPdf, pdfDesc.SupportedActionIds);
        Assert.Contains(ActionIds.EditPdfMetadata, pdfDesc.SupportedActionIds);
    }
}
