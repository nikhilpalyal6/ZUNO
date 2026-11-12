using System.IO;
using SmartShift.Core.IO;
using SmartShift.Processing;

namespace SmartShift.Processing.Documents;

public sealed class OfficeConversionEngine : IOfficeConversionEngine
{
    private readonly IOutputPathAllocator _allocator;

    public OfficeConversionEngine(IOutputPathAllocator allocator)
    {
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
    }

    public async Task<OfficeConversionResult> ConvertDocxToPdfAsync(
        string docxPath,
        string? outputPdfPath = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(docxPath) || !File.Exists(docxPath))
        {
            throw new FileNotFoundException("Source Word document does not exist.", docxPath);
        }

        var sourceHashBefore = FileHasher.Sha256Hex(docxPath);

        if (string.IsNullOrWhiteSpace(outputPdfPath))
        {
            _allocator.EnsureRoot();
            var baseName = $"{Path.GetFileNameWithoutExtension(docxPath)}.pdf";
            outputPdfPath = _allocator.Allocate(docxPath, baseName);
        }

        var result = await Task.Run(
            () => DocxToPdfConverter.ConvertWithReport(docxPath, outputPdfPath, null, progress, cancellationToken),
            cancellationToken).ConfigureAwait(false);

        var sourceHashAfter = FileHasher.Sha256Hex(docxPath);
        if (!string.Equals(sourceHashBefore, sourceHashAfter, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Source file immutability violation: source DOCX was modified during PDF conversion.");
        }

        return result;
    }

    public async Task<OfficeConversionResult> ConvertPptxToPdfAsync(
        string pptxPath,
        string slideRange = "all",
        string? outputPdfPath = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(pptxPath) || !File.Exists(pptxPath))
        {
            throw new FileNotFoundException("Source PowerPoint presentation does not exist.", pptxPath);
        }

        var sourceHashBefore = FileHasher.Sha256Hex(pptxPath);

        if (string.IsNullOrWhiteSpace(outputPdfPath))
        {
            _allocator.EnsureRoot();
            var baseName = $"{Path.GetFileNameWithoutExtension(pptxPath)}.pdf";
            outputPdfPath = _allocator.Allocate(pptxPath, baseName);
        }

        var result = await Task.Run(
            () => PptxToPdfConverter.Convert(pptxPath, outputPdfPath, slideRange, progress, cancellationToken),
            cancellationToken).ConfigureAwait(false);

        var sourceHashAfter = FileHasher.Sha256Hex(pptxPath);
        if (!string.Equals(sourceHashBefore, sourceHashAfter, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Source file immutability violation: source PPTX was modified during PDF conversion.");
        }

        return result;
    }

    public async Task<OfficeMultiOutputResult> ExportPptxSlidesAsync(
        string pptxPath,
        SlideExportFormat format = SlideExportFormat.Png,
        int dpi = 150,
        int quality = 85,
        string slideRange = "all",
        string? outputDirectory = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(pptxPath) || !File.Exists(pptxPath))
        {
            throw new FileNotFoundException("Source PowerPoint presentation does not exist.", pptxPath);
        }

        var sourceHashBefore = FileHasher.Sha256Hex(pptxPath);

        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            _allocator.EnsureRoot();
            outputDirectory = _allocator.RootDirectory;
        }

        var result = await Task.Run(
            () => PptxSlideImageExporter.Export(pptxPath, format, dpi, quality, slideRange, outputDirectory, progress, cancellationToken),
            cancellationToken).ConfigureAwait(false);

        var sourceHashAfter = FileHasher.Sha256Hex(pptxPath);
        if (!string.Equals(sourceHashBefore, sourceHashAfter, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Source file immutability violation: source PPTX was modified during slide export.");
        }

        return result;
    }

    public async Task<OfficeTextResult> ExtractDocxTextAsync(
        string docxPath,
        string? outputTxtPath = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(docxPath) || !File.Exists(docxPath))
        {
            throw new FileNotFoundException("Source Word document does not exist.", docxPath);
        }

        var sourceHashBefore = FileHasher.Sha256Hex(docxPath);

        if (string.IsNullOrWhiteSpace(outputTxtPath))
        {
            _allocator.EnsureRoot();
            var baseName = $"{Path.GetFileNameWithoutExtension(docxPath)}.txt";
            outputTxtPath = _allocator.Allocate(docxPath, baseName);
        }

        var result = await Task.Run(
            () => DocxTextExtractor.Extract(docxPath, outputTxtPath, progress, cancellationToken),
            cancellationToken).ConfigureAwait(false);

        var sourceHashAfter = FileHasher.Sha256Hex(docxPath);
        if (!string.Equals(sourceHashBefore, sourceHashAfter, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Source file immutability violation: source DOCX was modified during text extraction.");
        }

        return result;
    }

    public async Task<OfficeTextResult> ExtractPptxTextAsync(
        string pptxPath,
        string slideRange = "all",
        string? outputTxtPath = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(pptxPath) || !File.Exists(pptxPath))
        {
            throw new FileNotFoundException("Source PowerPoint presentation does not exist.", pptxPath);
        }

        var sourceHashBefore = FileHasher.Sha256Hex(pptxPath);

        if (string.IsNullOrWhiteSpace(outputTxtPath))
        {
            _allocator.EnsureRoot();
            var baseName = $"{Path.GetFileNameWithoutExtension(pptxPath)}.txt";
            outputTxtPath = _allocator.Allocate(pptxPath, baseName);
        }

        var result = await Task.Run(
            () => PptxTextExtractor.Extract(pptxPath, slideRange, outputTxtPath, progress, cancellationToken),
            cancellationToken).ConfigureAwait(false);

        var sourceHashAfter = FileHasher.Sha256Hex(pptxPath);
        if (!string.Equals(sourceHashBefore, sourceHashAfter, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Source file immutability violation: source PPTX was modified during text extraction.");
        }

        return result;
    }

    public async Task<DocxHtmlResult> ConvertDocxToHtmlAsync(
        string docxPath,
        string? outputHtmlPath = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(docxPath) || !File.Exists(docxPath))
        {
            throw new FileNotFoundException("Source Word document does not exist.", docxPath);
        }

        var sourceHashBefore = FileHasher.Sha256Hex(docxPath);

        if (string.IsNullOrWhiteSpace(outputHtmlPath))
        {
            _allocator.EnsureRoot();
            var baseName = $"{Path.GetFileNameWithoutExtension(docxPath)}.html";
            outputHtmlPath = _allocator.Allocate(docxPath, baseName);
        }

        var result = await Task.Run(
            () => DocxToHtmlConverter.Convert(docxPath, outputHtmlPath, progress, cancellationToken),
            cancellationToken).ConfigureAwait(false);

        var sourceHashAfter = FileHasher.Sha256Hex(docxPath);
        if (!string.Equals(sourceHashBefore, sourceHashAfter, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Source file immutability violation: source DOCX was modified during HTML conversion.");
        }

        return result;
    }
}
