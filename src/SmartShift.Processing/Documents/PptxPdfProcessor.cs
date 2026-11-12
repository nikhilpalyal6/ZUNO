using SmartShift.Core.Actions;
using SmartShift.Core.Classification;
using SmartShift.Core.IO;
using SmartShift.Processing;

namespace SmartShift.Processing.Documents;

public sealed class PptxPdfProcessor : IFileProcessor
{
    private readonly IOutputPathAllocator _allocator;
    private readonly FileClassifier _classifier;

    public PptxPdfProcessor(IOutputPathAllocator allocator, FileClassifier classifier)
    {
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
        _classifier = classifier ?? throw new ArgumentNullException(nameof(classifier));
    }

    public string ActionId => ActionIds.ConvertPptxToPdf;

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
            var baseName = request.Options.Values.TryGetValue("outputFileName", out var customName) && !string.IsNullOrWhiteSpace(customName)
                ? customName.Trim()
                : $"{Path.GetFileNameWithoutExtension(request.Input.Path)}.pdf";

            if (!baseName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                baseName += ".pdf";
            }

            var slideRange = request.Options.Values.TryGetValue("slideMode", out var sr) && !string.IsNullOrWhiteSpace(sr)
                ? sr.Trim()
                : (request.Options.Values.TryGetValue("pageRange", out var pr) && !string.IsNullOrWhiteSpace(pr) ? pr.Trim() : "all");

            var finalPath = _allocator.Allocate(request.Input.Path, baseName);

            var result = await Task.Run(
                () => PptxToPdfConverter.Convert(request.Input.Path, finalPath, slideRange, progress, cancellationToken),
                cancellationToken).ConfigureAwait(false);

            var details = new Dictionary<string, string>
            {
                ["OutputPath"] = result.OutputPath,
                ["outputPaths"] = result.OutputPath,
                ["qualityStatus"] = result.QualityReport.Status.ToString(),
                ["slidesCount"] = result.QualityReport.PagesOrSlidesCount.ToString(),
                ["headingsDetected"] = result.QualityReport.HeadingsDetectedCount.ToString(),
                ["tablesReconstructed"] = result.QualityReport.TablesReconstructedCount.ToString(),
                ["imagesPreserved"] = result.QualityReport.ImagesPreservedCount.ToString(),
                ["shapesPreserved"] = result.QualityReport.ShapesPreservedCount.ToString(),
                ["conversionMode"] = result.QualityReport.ConversionMode
            };

            var allNotes = result.QualityReport.Warnings.Concat(result.QualityReport.Limitations).ToList();
            if (allNotes.Count > 0)
            {
                details["warnings"] = string.Join("; ", allNotes);
            }

            return ProcessingResult.Success(
                result.SummaryNote,
                result.OutputPath,
                details);
        }
        catch (OperationCanceledException)
        {
            return ProcessingResult.Cancelled();
        }
        catch (Exception ex)
        {
            return ProcessingResult.Fail($"PowerPoint to PDF conversion failed: {ex.Message}");
        }
    }
}
