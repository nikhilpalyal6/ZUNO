namespace SmartShift.Processing.Documents;

public interface IOfficeConversionEngine
{
    Task<OfficeConversionResult> ConvertDocxToPdfAsync(
        string docxPath,
        string? outputPdfPath = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);

    Task<OfficeConversionResult> ConvertPptxToPdfAsync(
        string pptxPath,
        string slideRange = "all",
        string? outputPdfPath = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);

    Task<OfficeMultiOutputResult> ExportPptxSlidesAsync(
        string pptxPath,
        SlideExportFormat format = SlideExportFormat.Png,
        int dpi = 150,
        int quality = 85,
        string slideRange = "all",
        string? outputDirectory = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);

    Task<OfficeTextResult> ExtractDocxTextAsync(
        string docxPath,
        string? outputTxtPath = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);

    Task<OfficeTextResult> ExtractPptxTextAsync(
        string pptxPath,
        string slideRange = "all",
        string? outputTxtPath = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);

    Task<DocxHtmlResult> ConvertDocxToHtmlAsync(
        string docxPath,
        string? outputHtmlPath = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default);
}
