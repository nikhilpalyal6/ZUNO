using SmartShift.Processing.Documents.Editing;

namespace SmartShift.Processing.Pdf.Editing;

public interface IPdfWatermarkService
{
    Task<DocumentEditingResult> AddWatermarkAsync(
        string inputPath,
        string outputPath,
        PdfWatermarkOptions options,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public interface IPdfMetadataService
{
    Task<PdfMetadata> ReadMetadataAsync(
        string inputPath,
        CancellationToken cancellationToken = default);

    Task<DocumentEditingResult> UpdateMetadataAsync(
        string inputPath,
        string outputPath,
        PdfMetadata metadata,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);
}
