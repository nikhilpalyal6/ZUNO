using SmartShift.Core.Actions;
using SmartShift.Core.Classification;
using SmartShift.Core.IO;
using SmartShift.Processing;

namespace SmartShift.Processing.Documents;

public sealed class PptxExportProcessor : IFileProcessor
{
    private readonly IOutputPathAllocator _allocator;
    private readonly FileClassifier _classifier;

    public PptxExportProcessor(IOutputPathAllocator allocator, FileClassifier classifier)
    {
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
        _classifier = classifier ?? throw new ArgumentNullException(nameof(classifier));
    }

    public string ActionId => ActionIds.ExportPptxImages;

    public bool IsAvailable => true;

    public async Task<ProcessingResult> ProcessAsync(
        ProcessingRequest request,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(request.Input.Path) || !File.Exists(request.Input.Path))
        {
            return ProcessingResult.Fail("The source PowerPoint presentation does not exist or has been moved.");
        }

        try
        {
            _allocator.EnsureRoot();
            var outDir = string.IsNullOrWhiteSpace(request.OutputDirectory)
                ? _allocator.RootDirectory
                : request.OutputDirectory;

            var rawFmt = request.Options.Values.TryGetValue("format", out var f) && !string.IsNullOrWhiteSpace(f)
                ? f.Trim().ToLowerInvariant()
                : "png";

            var format = rawFmt switch
            {
                "jpeg" or "jpg" => SlideExportFormat.Jpeg,
                "webp" => SlideExportFormat.Webp,
                _ => SlideExportFormat.Png
            };

            int dpi = 150;
            if (request.Options.Values.TryGetValue("dpi", out var rawDpi) && int.TryParse(rawDpi, out int parsedDpi))
            {
                dpi = parsedDpi;
            }

            int quality = 85;
            if (request.Options.Values.TryGetValue("quality", out var rawQ) && int.TryParse(rawQ, out int parsedQ))
            {
                quality = parsedQ;
            }

            var slideRange = request.Options.Values.TryGetValue("pageRange", out var pr) && !string.IsNullOrWhiteSpace(pr)
                ? pr.Trim()
                : (request.Options.Values.TryGetValue("slideRange", out var sr) && !string.IsNullOrWhiteSpace(sr) ? sr.Trim() : "all");

            var result = await Task.Run(
                () => PptxSlideImageExporter.Export(request.Input.Path, format, dpi, quality, slideRange, outDir, progress, cancellationToken),
                cancellationToken).ConfigureAwait(false);

            if (result.OutputPaths.Count == 0)
            {
                return ProcessingResult.Fail("No slide images were generated.");
            }

            var primaryOutput = result.OutputPaths[0];
            var details = new Dictionary<string, string>
            {
                ["OutputPath"] = primaryOutput,
                ["outputPaths"] = string.Join(";", result.OutputPaths),
                ["totalExported"] = result.OutputPaths.Count.ToString(),
                ["format"] = format.ToString().ToLowerInvariant(),
                ["dpi"] = dpi.ToString(),
                ["qualityStatus"] = result.QualityReport.Status.ToString(),
                ["conversionMode"] = result.QualityReport.ConversionMode
            };

            var allNotes = result.QualityReport.Warnings.Concat(result.QualityReport.Limitations).ToList();
            if (allNotes.Count > 0)
            {
                details["warnings"] = string.Join("; ", allNotes);
            }

            return ProcessingResult.Success(
                result.SummaryNote,
                primaryOutput,
                details);
        }
        catch (OperationCanceledException)
        {
            return ProcessingResult.Cancelled();
        }
        catch (Exception ex)
        {
            return ProcessingResult.Fail($"PowerPoint slide export failed: {ex.Message}");
        }
    }
}
