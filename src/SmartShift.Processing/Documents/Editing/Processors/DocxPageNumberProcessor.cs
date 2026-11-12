using SmartShift.Core.Actions;
using SmartShift.Core.Classification;
using SmartShift.Core.IO;
using SmartShift.Processing;
using SmartShift.Processing.Documents.Editing;

namespace SmartShift.Processing.Documents.Editing.Processors;

public sealed class DocxPageNumberProcessor : IFileProcessor
{
    private readonly IDocxPageNumberService _pageNumberService;
    private readonly IOutputPathAllocator _allocator;
    private readonly FileClassifier _classifier;

    public DocxPageNumberProcessor(IDocxPageNumberService pageNumberService, IOutputPathAllocator allocator, FileClassifier classifier)
    {
        _pageNumberService = pageNumberService ?? throw new ArgumentNullException(nameof(pageNumberService));
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
        _classifier = classifier ?? throw new ArgumentNullException(nameof(classifier));
    }

    public string ActionId => ActionIds.AddDocxPageNumbers;

    public bool IsAvailable => true;

    public async Task<ProcessingResult> ProcessAsync(
        ProcessingRequest request,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(request.Input.Path) || !File.Exists(request.Input.Path))
        {
            return ProcessingResult.Fail("The source Word document does not exist or has been moved.");
        }

        try
        {
            _allocator.EnsureRoot();
            var baseName = (request.Options.Values.TryGetValue("outputFileName", out var customName) ||
                            request.Options.Values.TryGetValue("customOutputName", out customName)) &&
                           !string.IsNullOrWhiteSpace(customName)
                ? customName.Trim()
                : $"{Path.GetFileNameWithoutExtension(request.Input.Path)}_numbered.docx";

            if (!baseName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                baseName += ".docx";
            }

            var position = PageNumberPosition.Footer;
            if (request.Options.Values.TryGetValue("position", out var posStr) &&
                string.Equals(posStr, "header", StringComparison.OrdinalIgnoreCase))
            {
                position = PageNumberPosition.Header;
            }

            var alignment = PageNumberAlignment.Center;
            if (request.Options.Values.TryGetValue("alignment", out var alignStr))
            {
                if (string.Equals(alignStr, "left", StringComparison.OrdinalIgnoreCase)) alignment = PageNumberAlignment.Left;
                else if (string.Equals(alignStr, "right", StringComparison.OrdinalIgnoreCase)) alignment = PageNumberAlignment.Right;
            }

            var format = PageNumberFormat.PageNumberTotal;
            if (request.Options.Values.TryGetValue("format", out var formatStr))
            {
                if (string.Equals(formatStr, "numberOnly", StringComparison.OrdinalIgnoreCase) || formatStr == "1")
                    format = PageNumberFormat.NumberOnly;
                else if (string.Equals(formatStr, "pageNumber", StringComparison.OrdinalIgnoreCase))
                    format = PageNumberFormat.PageNumber;
            }

            var options = new DocxPageNumberOptions(
                Position: position,
                Alignment: alignment,
                Format: format);

            var finalPath = _allocator.Allocate(request.Input.Path, baseName);
            var result = await _pageNumberService.AddPageNumbersAsync(request.Input.Path, finalPath, options, progress, cancellationToken)
                .ConfigureAwait(false);

            var details = new Dictionary<string, string>
            {
                ["OutputPath"] = result.OutputPath,
                ["outputPaths"] = result.OutputPath,
                ["Position"] = position.ToString(),
                ["Alignment"] = alignment.ToString(),
                ["Format"] = format.ToString()
            };

            return ProcessingResult.Success(
                result.Message ?? "Page numbers added successfully.",
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
