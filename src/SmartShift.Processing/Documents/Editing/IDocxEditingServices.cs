namespace SmartShift.Processing.Documents.Editing;

public interface IDocxMergeService
{
    Task<DocxMergeResult> MergeAsync(
        IReadOnlyList<string> inputPaths,
        string outputPath,
        DocxMergeOptions? options = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public interface IDocxPageNumberService
{
    Task<DocumentEditingResult> AddPageNumbersAsync(
        string inputPath,
        string outputPath,
        DocxPageNumberOptions? options = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public interface IDocxWatermarkService
{
    Task<DocumentEditingResult> AddWatermarkAsync(
        string inputPath,
        string outputPath,
        DocxWatermarkOptions options,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public interface IDocxMetadataService
{
    Task<DocxMetadata> ReadMetadataAsync(
        string inputPath,
        CancellationToken cancellationToken = default);

    Task<DocumentEditingResult> UpdateMetadataAsync(
        string inputPath,
        string outputPath,
        DocxMetadata metadata,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);
}
