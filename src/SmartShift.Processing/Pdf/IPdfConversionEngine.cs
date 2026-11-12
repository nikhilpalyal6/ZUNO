using SmartShift.Processing.Pdf.Office;

namespace SmartShift.Processing.Pdf;

/// <summary>
/// Production engine coordinating high-fidelity PDF conversions to Office formats, image formats, and text.
/// </summary>
public interface IPdfConversionEngine
{
    Task<PdfDocxResult> ConvertToDocxAsync(
        string inputPath,
        string? pageRange,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken);

    Task<PdfPptxResult> ConvertToPptxAsync(
        string inputPath,
        PptxSlideMode mode,
        string? pageRange,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken);

    Task<PdfExportResult> ExportPagesAsync(
        string inputPath,
        ExportImageFormat format,
        int dpi,
        int quality,
        string? pageRange,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken);

    Task<PdfTextExtractionResult> ExtractTextAsync(
        string inputPath,
        string? pageRange,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken);
}
