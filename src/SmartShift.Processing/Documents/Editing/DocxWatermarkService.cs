using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Vml = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using SmartShift.Core.IO;
using SmartShift.Processing;

namespace SmartShift.Processing.Documents.Editing;

public sealed class DocxWatermarkService : IDocxWatermarkService
{
    private readonly IOutputPathAllocator _allocator;

    public DocxWatermarkService(IOutputPathAllocator allocator)
    {
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
    }

    public async Task<DocumentEditingResult> AddWatermarkAsync(
        string inputPath,
        string outputPath,
        DocxWatermarkOptions options,
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

        ArgumentNullException.ThrowIfNull(options);
        if (string.IsNullOrWhiteSpace(options.Text))
        {
            throw new ArgumentException("Watermark text cannot be empty.", nameof(options));
        }

        cancellationToken.ThrowIfCancellationRequested();

        return await Task.Run(() => AddWatermarkInternal(inputPath, outputPath, options, progress, cancellationToken), cancellationToken)
            .ConfigureAwait(false);
    }

    private static DocumentEditingResult AddWatermarkInternal(
        string inputPath,
        string outputPath,
        DocxWatermarkOptions options,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        progress?.Report(new ProcessingProgress(10, "Verifying input document…"));
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

        try
        {
            progress?.Report(new ProcessingProgress(25, "Preparing watermark structure…"));
            File.Copy(inputPath, tempPartial, overwrite: true);

            using (var doc = WordprocessingDocument.Open(tempPartial, true))
            {
                var mainPart = doc.MainDocumentPart
                    ?? throw new InvalidOperationException("The Word document does not contain a valid MainDocumentPart.");

                if (mainPart.Document == null)
                {
                    mainPart.Document = new W.Document();
                }

                var body = mainPart.Document.Body ?? mainPart.Document.AppendChild(new W.Body());

                progress?.Report(new ProcessingProgress(50, "Embedding watermark in document headers…"));

                var hexColor = options.HexColor.StartsWith('#') ? options.HexColor : "#" + options.HexColor;
                var rotation = options.IsDiagonal ? "315" : "0";
                var opacityStr = Math.Clamp(options.Opacity, 0.05, 1.0).ToString("0.00", CultureInfo.InvariantCulture);
                var fontSizePt = Math.Clamp(options.FontSize, 12, 144).ToString(CultureInfo.InvariantCulture);

                var watermarkParagraph = CreateWatermarkParagraph(options.Text, options.FontName, fontSizePt, hexColor, rotation, opacityStr);

                // Add or update HeaderPart
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                headerPart.Header = new W.Header(watermarkParagraph);
                headerPart.Header.Save();
                var relId = mainPart.GetIdOfPart(headerPart);

                // Apply to sections
                var sectionList = body.Descendants<W.SectionProperties>().ToList();
                var bodySectPr = body.Elements<W.SectionProperties>().LastOrDefault();
                if (bodySectPr != null && !sectionList.Contains(bodySectPr))
                {
                    sectionList.Add(bodySectPr);
                }

                if (sectionList.Count == 0)
                {
                    var newSectPr = new W.SectionProperties(new W.HeaderReference { Type = W.HeaderFooterValues.Default, Id = relId });
                    body.AppendChild(newSectPr);
                }
                else
                {
                    foreach (var sectPr in sectionList)
                    {
                        // Remove existing watermark headers if present
                        var existingHeaders = sectPr.Elements<W.HeaderReference>()
                            .Where(h => h.Type?.Value == W.HeaderFooterValues.Default)
                            .ToList();

                        foreach (var eh in existingHeaders)
                        {
                            eh.Remove();
                        }

                        sectPr.PrependChild(new W.HeaderReference { Type = W.HeaderFooterValues.Default, Id = relId });
                    }
                }

                mainPart.Document.Save();
            }

            progress?.Report(new ProcessingProgress(90, "Finalizing watermarked document…"));

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
            File.Move(tempPartial, outputPath, overwrite: true);

            var postHash = ComputeSha256(inputPath);
            if (!string.Equals(preHash, postHash, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"CRITICAL: Source file '{inputPath}' was modified during watermark operation!");
            }

            progress?.Report(new ProcessingProgress(100, "Watermark successfully embedded."));

            return new DocumentEditingResult(
                OutputPath: outputPath,
                IsSuccess: true,
                Limitations: [],
                Message: $"Watermark '{options.Text}' successfully embedded in document header.");
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

    private static W.Paragraph CreateWatermarkParagraph(
        string text,
        string fontName,
        string fontSizePt,
        string hexColor,
        string rotation,
        string opacityStr)
    {
        var shape = new Vml.Shape
        {
            Id = "WordWatermark",
            Style = $"position:absolute;margin-left:0;margin-top:0;width:410pt;height:100pt;rotation:{rotation};z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
            FillColor = hexColor,
            Stroked = false
        };

        shape.AppendChild(new Vml.Fill { Opacity = opacityStr });
        shape.AppendChild(new Vml.TextPath
        {
            Style = $"font-family:'{fontName}';font-size:{fontSizePt}pt",
            String = text
        });

        var pict = new W.Picture(shape);
        var run = new W.Run(pict);
        return new W.Paragraph(run);
    }

    private static string ComputeSha256(string path)
    {
        using var sha = SHA256.Create();
        using var stream = File.OpenRead(path);
        var bytes = sha.ComputeHash(stream);
        return Convert.ToHexString(bytes);
    }
}
