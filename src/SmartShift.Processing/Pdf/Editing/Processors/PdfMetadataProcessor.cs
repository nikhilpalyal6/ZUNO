using SmartShift.Core.Actions;
using SmartShift.Core.Classification;
using SmartShift.Core.IO;
using SmartShift.Processing;
using SmartShift.Processing.Pdf.Editing;

namespace SmartShift.Processing.Pdf.Editing.Processors;

public sealed class PdfMetadataProcessor : IFileProcessor
{
    private readonly IPdfMetadataService _metadataService;
    private readonly IOutputPathAllocator _allocator;
    private readonly FileClassifier _classifier;

    public PdfMetadataProcessor(IPdfMetadataService metadataService, IOutputPathAllocator allocator, FileClassifier classifier)
    {
        _metadataService = metadataService ?? throw new ArgumentNullException(nameof(metadataService));
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
        _classifier = classifier ?? throw new ArgumentNullException(nameof(classifier));
    }

    public string ActionId => ActionIds.EditPdfMetadata;

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
                : $"{Path.GetFileNameWithoutExtension(request.Input.Path)}_meta.pdf";

            if (!baseName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                baseName += ".pdf";
            }

            request.Options.Values.TryGetValue("title", out var title);
            request.Options.Values.TryGetValue("author", out var author);
            request.Options.Values.TryGetValue("subject", out var subject);
            request.Options.Values.TryGetValue("keywords", out var keywords);
            request.Options.Values.TryGetValue("creator", out var creator);
            request.Options.Values.TryGetValue("producer", out var producer);

            var metadata = new PdfMetadata(
                Title: title,
                Author: author,
                Subject: subject,
                Keywords: keywords,
                Creator: creator,
                Producer: producer);

            var finalPath = _allocator.Allocate(request.Input.Path, baseName);
            var result = await _metadataService.UpdateMetadataAsync(request.Input.Path, finalPath, metadata, progress, cancellationToken)
                .ConfigureAwait(false);

            var details = new Dictionary<string, string>
            {
                ["OutputPath"] = result.OutputPath,
                ["outputPaths"] = result.OutputPath
            };

            if (title != null) details["Title"] = title;
            if (author != null) details["Author"] = author;

            return ProcessingResult.Success(
                result.Message ?? "PDF metadata updated successfully.",
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
