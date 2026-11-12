using System.Globalization;
using SmartShift.Core.Actions;
using SmartShift.Core.Classification;
using SmartShift.Core.IO;
using SmartShift.Processing;
using SmartShift.Processing.Pdf.Editing;

namespace SmartShift.Processing.Pdf.Editing.Processors;

public sealed class PdfWatermarkProcessor : IFileProcessor
{
    private readonly IPdfWatermarkService _watermarkService;
    private readonly IOutputPathAllocator _allocator;
    private readonly FileClassifier _classifier;

    public PdfWatermarkProcessor(IPdfWatermarkService watermarkService, IOutputPathAllocator allocator, FileClassifier classifier)
    {
        _watermarkService = watermarkService ?? throw new ArgumentNullException(nameof(watermarkService));
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
        _classifier = classifier ?? throw new ArgumentNullException(nameof(classifier));
    }

    public string ActionId => ActionIds.WatermarkPdf;

    public bool IsAvailable => true;

    public async Task<ProcessingResult> ProcessAsync(
        ProcessingRequest request,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(request.Input.Path) || !File.Exists(request.Input.Path))
        {
            return ProcessingResult.Fail("The source PDF document does not exist or has been moved.");
        }

        try
        {
            _allocator.EnsureRoot();
            var baseName = (request.Options.Values.TryGetValue("outputFileName", out var customName) ||
                            request.Options.Values.TryGetValue("customOutputName", out customName)) &&
                           !string.IsNullOrWhiteSpace(customName)
                ? customName.Trim()
                : $"{Path.GetFileNameWithoutExtension(request.Input.Path)}_watermarked.pdf";

            if (!baseName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                baseName += ".pdf";
            }

            var text = (request.Options.Values.TryGetValue("text", out var t) ||
                        request.Options.Values.TryGetValue("watermarkText", out t)) &&
                       !string.IsNullOrWhiteSpace(t)
                ? t.Trim()
                : "CONFIDENTIAL";

            double fontSize = 48.0;
            if (request.Options.Values.TryGetValue("fontSize", out var fsStr) &&
                double.TryParse(fsStr, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsedFs))
            {
                fontSize = parsedFs;
            }

            double opacity = 0.25;
            if (request.Options.Values.TryGetValue("opacity", out var opStr) &&
                double.TryParse(opStr, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsedOp))
            {
                opacity = parsedOp;
            }

            bool isDiagonal = true;
            if (request.Options.Values.TryGetValue("diagonal", out var diagStr))
            {
                if (string.Equals(diagStr, "false", StringComparison.OrdinalIgnoreCase) || diagStr == "0")
                {
                    isDiagonal = false;
                }
            }

            var color = (request.Options.Values.TryGetValue("color", out var c) ||
                         request.Options.Values.TryGetValue("hexColor", out c)) &&
                        !string.IsNullOrWhiteSpace(c)
                ? c.Trim()
                : "808080";

            var fontName = request.Options.Values.TryGetValue("fontName", out var fn) && !string.IsNullOrWhiteSpace(fn)
                ? fn.Trim()
                : "Arial";

            var pageRange = (request.Options.Values.TryGetValue("pageRange", out var pr) ||
                             request.Options.Values.TryGetValue("pages", out pr)) &&
                            !string.IsNullOrWhiteSpace(pr)
                ? pr.Trim()
                : "all";

            var options = new PdfWatermarkOptions(
                Text: text,
                FontSize: fontSize,
                Opacity: opacity,
                IsDiagonal: isDiagonal,
                HexColor: color,
                FontName: fontName,
                PageRange: pageRange);

            var finalPath = _allocator.Allocate(request.Input.Path, baseName);
            var result = await _watermarkService.AddWatermarkAsync(request.Input.Path, finalPath, options, progress, cancellationToken)
                .ConfigureAwait(false);

            var details = new Dictionary<string, string>
            {
                ["OutputPath"] = result.OutputPath,
                ["outputPaths"] = result.OutputPath,
                ["WatermarkText"] = text,
                ["FontSize"] = fontSize.ToString(CultureInfo.InvariantCulture),
                ["Opacity"] = opacity.ToString(CultureInfo.InvariantCulture),
                ["Diagonal"] = isDiagonal.ToString(),
                ["PageRange"] = pageRange
            };

            return ProcessingResult.Success(
                result.Message ?? $"Watermark '{text}' applied successfully to PDF.",
                result.OutputPath,
                details);
        }
        catch (OperationCanceledException)
        {
            return ProcessingResult.Cancelled();
        }
        catch (Exception ex)
        {
            return ProcessingResult.Fail(ex.Message);
        }
    }
}
