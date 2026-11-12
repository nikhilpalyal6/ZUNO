using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SmartShift.Core.IO;
using SmartShift.Processing;

namespace SmartShift.Processing.Documents.Editing;

public sealed class DocxPageNumberService : IDocxPageNumberService
{
    private readonly IOutputPathAllocator _allocator;

    public DocxPageNumberService(IOutputPathAllocator allocator)
    {
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
    }

    public async Task<DocumentEditingResult> AddPageNumbersAsync(
        string inputPath,
        string outputPath,
        DocxPageNumberOptions? options = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
        {
            throw new FileNotFoundException($"Input DOCX file not found: {inputPath}", inputPath);
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            throw new ArgumentException("Output path must be specified.", nameof(outputPath));
        }

        options ??= new DocxPageNumberOptions();
        cancellationToken.ThrowIfCancellationRequested();

        return await Task.Run(() => AddPageNumbersInternal(inputPath, outputPath, options, progress, cancellationToken), cancellationToken)
            .ConfigureAwait(false);
    }

    private static DocumentEditingResult AddPageNumbersInternal(
        string inputPath,
        string outputPath,
        DocxPageNumberOptions options,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        progress?.Report(new ProcessingProgress(10, "Verifying document integrity…"));
        var preHash = ComputeSha256(inputPath);

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

        var limitations = new List<string>();

        try
        {
            progress?.Report(new ProcessingProgress(25, "Preparing document for page numbering…"));
            File.Copy(inputPath, tempPartial, overwrite: true);

            using (var doc = WordprocessingDocument.Open(tempPartial, true))
            {
                var mainPart = doc.MainDocumentPart
                    ?? throw new InvalidOperationException("The Word document does not contain a valid MainDocumentPart.");

                if (mainPart.Document == null)
                {
                    mainPart.Document = new Document();
                }

                var body = mainPart.Document.Body ?? mainPart.Document.AppendChild(new Body());

                progress?.Report(new ProcessingProgress(50, "Generating header/footer page number fields…"));

                var justification = options.Alignment switch
                {
                    PageNumberAlignment.Left => JustificationValues.Left,
                    PageNumberAlignment.Right => JustificationValues.Right,
                    _ => JustificationValues.Center
                };

                var fontSizeHalfPoints = ((int)Math.Round(options.FontSize * 2)).ToString();
                var runProps = new RunProperties(
                    new RunFonts { Ascii = options.FontName, HighAnsi = options.FontName },
                    new FontSize { Val = fontSizeHalfPoints });

                var pPr = new ParagraphProperties(new Justification { Val = justification });
                var p = new Paragraph(pPr);

                switch (options.Format)
                {
                    case PageNumberFormat.NumberOnly:
                        p.AppendChild(new SimpleField { Instruction = "PAGE" });
                        break;

                    case PageNumberFormat.PageNumber:
                        p.AppendChild(new Run((RunProperties)runProps.CloneNode(true), new Text("Page ") { Space = SpaceProcessingModeValues.Preserve }));
                        p.AppendChild(new SimpleField { Instruction = "PAGE" });
                        break;

                    case PageNumberFormat.PageNumberTotal:
                    default:
                        p.AppendChild(new Run((RunProperties)runProps.CloneNode(true), new Text("Page ") { Space = SpaceProcessingModeValues.Preserve }));
                        p.AppendChild(new SimpleField { Instruction = "PAGE" });
                        p.AppendChild(new Run((RunProperties)runProps.CloneNode(true), new Text(" of ") { Space = SpaceProcessingModeValues.Preserve }));
                        p.AppendChild(new SimpleField { Instruction = "NUMPAGES" });
                        break;
                }

                if (options.Position == PageNumberPosition.Footer)
                {
                    var footerPart = mainPart.AddNewPart<FooterPart>();
                    footerPart.Footer = new Footer(p);
                    footerPart.Footer.Save();
                    var relId = mainPart.GetIdOfPart(footerPart);

                    ApplyFooterReferenceToSections(body, relId, options.ApplyToAllSections);
                }
                else
                {
                    var headerPart = mainPart.AddNewPart<HeaderPart>();
                    headerPart.Header = new Header(p);
                    headerPart.Header.Save();
                    var relId = mainPart.GetIdOfPart(headerPart);

                    ApplyHeaderReferenceToSections(body, relId, options.ApplyToAllSections);
                }

                mainPart.Document.Save();
            }

            progress?.Report(new ProcessingProgress(90, "Finalizing page numbered document…"));

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
            File.Move(tempPartial, outputPath, overwrite: true);

            var postHash = ComputeSha256(inputPath);
            if (!string.Equals(preHash, postHash, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"CRITICAL: Source file '{inputPath}' was modified during page numbering operation!");
            }

            progress?.Report(new ProcessingProgress(100, "Page numbers successfully added."));

            return new DocumentEditingResult(
                OutputPath: outputPath,
                IsSuccess: true,
                Limitations: limitations,
                Message: $"Page numbers ({options.Format}) added to {options.Position.ToString().ToLowerInvariant()} successfully.");
        }
        catch
        {
            if (File.Exists(tempPartial))
            {
                try { File.Delete(tempPartial); } catch { }
            }
            throw;
        }
    }

    private static void ApplyFooterReferenceToSections(Body body, string relId, bool applyToAllSections)
    {
        var sectionList = body.Descendants<SectionProperties>().ToList();
        var bodySectPr = body.Elements<SectionProperties>().LastOrDefault();
        if (bodySectPr != null && !sectionList.Contains(bodySectPr))
        {
            sectionList.Add(bodySectPr);
        }

        if (sectionList.Count == 0)
        {
            var newSectPr = new SectionProperties(new FooterReference { Type = HeaderFooterValues.Default, Id = relId });
            body.AppendChild(newSectPr);
            return;
        }

        var targets = applyToAllSections ? sectionList : [sectionList.Last()];
        foreach (var sectPr in targets)
        {
            var existingFooters = sectPr.Elements<FooterReference>()
                .Where(f => f.Type?.Value == HeaderFooterValues.Default)
                .ToList();

            foreach (var ef in existingFooters)
            {
                ef.Remove();
            }

            sectPr.PrependChild(new FooterReference { Type = HeaderFooterValues.Default, Id = relId });
        }
    }

    private static void ApplyHeaderReferenceToSections(Body body, string relId, bool applyToAllSections)
    {
        var sectionList = body.Descendants<SectionProperties>().ToList();
        var bodySectPr = body.Elements<SectionProperties>().LastOrDefault();
        if (bodySectPr != null && !sectionList.Contains(bodySectPr))
        {
            sectionList.Add(bodySectPr);
        }

        if (sectionList.Count == 0)
        {
            var newSectPr = new SectionProperties(new HeaderReference { Type = HeaderFooterValues.Default, Id = relId });
            body.AppendChild(newSectPr);
            return;
        }

        var targets = applyToAllSections ? sectionList : [sectionList.Last()];
        foreach (var sectPr in targets)
        {
            var existingHeaders = sectPr.Elements<HeaderReference>()
                .Where(h => h.Type?.Value == HeaderFooterValues.Default)
                .ToList();

            foreach (var eh in existingHeaders)
            {
                eh.Remove();
            }

            sectPr.PrependChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = relId });
        }
    }

    private static string ComputeSha256(string path)
    {
        using var sha = SHA256.Create();
        using var stream = File.OpenRead(path);
        var bytes = sha.ComputeHash(stream);
        return Convert.ToHexString(bytes);
    }
}
