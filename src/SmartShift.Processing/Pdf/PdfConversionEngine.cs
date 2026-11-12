using SmartShift.Core.IO;
using SmartShift.Processing.Pdf.Office;

namespace SmartShift.Processing.Pdf;

/// <summary>
/// Production conversion engine coordinating PDF transformations to Word (DOCX), PowerPoint (PPTX),
/// raster image sets (PNG, JPEG, WebP), and reading-order plain text.
/// </summary>
public sealed class PdfConversionEngine : IPdfConversionEngine
{
    private readonly IOutputPathAllocator _allocator;

    public PdfConversionEngine(IOutputPathAllocator allocator)
    {
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
    }

    public async Task<PdfDocxResult> ConvertToDocxAsync(
        string inputPath,
        string? pageRange,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        if (cancellationToken.IsCancellationRequested)
        {
            throw new OperationCanceledException(cancellationToken);
        }

        return await Task.Run(
            () => PdfToDocxConverter.ConvertToDocx(inputPath, _allocator, pageRange, progress, cancellationToken),
            cancellationToken).ConfigureAwait(false);
    }

    public async Task<PdfPptxResult> ConvertToPptxAsync(
        string inputPath,
        PptxSlideMode mode,
        string? pageRange,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        if (cancellationToken.IsCancellationRequested)
        {
            throw new OperationCanceledException(cancellationToken);
        }

        return await Task.Run(
            () => PdfToPptxConverter.ConvertToPptx(inputPath, _allocator, mode, pageRange, progress, cancellationToken),
            cancellationToken).ConfigureAwait(false);
    }

    public async Task<PdfExportResult> ExportPagesAsync(
        string inputPath,
        ExportImageFormat format,
        int dpi,
        int quality,
        string? pageRange,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        if (cancellationToken.IsCancellationRequested)
        {
            throw new OperationCanceledException(cancellationToken);
        }

        return await Task.Run(
            () => PdfPageExporter.ExportPages(inputPath, _allocator, format, dpi, quality, pageRange, progress, cancellationToken),
            cancellationToken).ConfigureAwait(false);
    }

    public async Task<PdfTextExtractionResult> ExtractTextAsync(
        string inputPath,
        string? pageRange,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        if (cancellationToken.IsCancellationRequested)
        {
            throw new OperationCanceledException(cancellationToken);
        }

        return await Task.Run(
            () => PdfTextExtractor.ExtractText(inputPath, _allocator, pageRange, progress, cancellationToken),
            cancellationToken).ConfigureAwait(false);
    }
}
