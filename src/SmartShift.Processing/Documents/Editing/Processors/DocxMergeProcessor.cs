using SmartShift.Core.Actions;
using SmartShift.Core.Classification;
using SmartShift.Core.IO;
using SmartShift.Processing;
using SmartShift.Processing.Documents.Editing;

namespace SmartShift.Processing.Documents.Editing.Processors;

public sealed class DocxMergeProcessor : IFileProcessor
{
    private readonly IDocxMergeService _mergeService;
    private readonly IOutputPathAllocator _allocator;
    private readonly FileClassifier _classifier;

    public DocxMergeProcessor(IDocxMergeService mergeService, IOutputPathAllocator allocator, FileClassifier classifier)
    {
        _mergeService = mergeService ?? throw new ArgumentNullException(nameof(mergeService));
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
        _classifier = classifier ?? throw new ArgumentNullException(nameof(classifier));
    }

    public string ActionId => ActionIds.MergeDocx;

    public bool IsAvailable => true;

    public async Task<ProcessingResult> ProcessAsync(
        ProcessingRequest request,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var files = new List<string>();
        if (!string.IsNullOrWhiteSpace(request.Input.Path))
        {
            files.Add(request.Input.Path);
        }

        if ((request.Options.Values.TryGetValue("additionalFiles", out var additional) ||
             request.Options.Values.TryGetValue("additionalPaths", out additional) ||
             request.Options.Values.TryGetValue("orderedFiles", out additional)) &&
            !string.IsNullOrWhiteSpace(additional))
        {
            var split = additional.Split([';', '|', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            files.AddRange(split);
        }

        if (files.Count < 2)
        {
            return ProcessingResult.Fail("At least 2 DOCX files are required to merge.");
        }

        try
        {
            _allocator.EnsureRoot();
            var baseName = (request.Options.Values.TryGetValue("outputFileName", out var customName) ||
                            request.Options.Values.TryGetValue("customOutputName", out customName)) &&
                           !string.IsNullOrWhiteSpace(customName)
                ? customName.Trim()
                : $"{Path.GetFileNameWithoutExtension(request.Input.Path)}_merged.docx";

            if (!baseName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                baseName += ".docx";
            }

            var finalPath = _allocator.Allocate(request.Input.Path, baseName);
            var result = await _mergeService.MergeAsync(files, finalPath, new DocxMergeOptions(), progress, cancellationToken)
                .ConfigureAwait(false);

            var details = new Dictionary<string, string>
            {
                ["OutputPath"] = result.OutputPath,
                ["outputPaths"] = result.OutputPath,
                ["TotalFiles"] = result.DocumentCount.ToString(),
                ["ParagraphCount"] = result.ParagraphCount.ToString(),
                ["TableCount"] = result.TableCount.ToString(),
                ["ImageCount"] = result.ImageCount.ToString()
            };

            return ProcessingResult.Success(
                $"Successfully merged {result.DocumentCount} DOCX documents ({result.ParagraphCount} paragraphs, {result.TableCount} tables).",
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
