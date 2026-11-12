namespace SmartShift.Processing.Documents;

public enum OfficeConversionStatus
{
    Successful = 0,
    SuccessfulWithLimitations = 1,
    Failed = 2,
    Cancelled = 3
}

public enum SlideExportFormat
{
    Png = 0,
    Jpeg = 1,
    Webp = 2
}

public sealed record OfficeConversionQualityReport(
    OfficeConversionStatus Status,
    string ConversionMode,
    int PagesOrSlidesCount,
    int HeadingsDetectedCount,
    int TablesReconstructedCount,
    int ImagesPreservedCount,
    int ShapesPreservedCount,
    long CharacterCount,
    IReadOnlyList<string> Limitations,
    IReadOnlyList<string> Warnings,
    bool ValidationPassed,
    TimeSpan Duration);

public sealed record OfficeConversionResult(
    string OutputPath,
    OfficeConversionQualityReport QualityReport,
    string SummaryNote);

public sealed record OfficeMultiOutputResult(
    IReadOnlyList<string> OutputPaths,
    OfficeConversionQualityReport QualityReport,
    string SummaryNote);

public sealed record OfficeTextResult(
    string OutputPath,
    string ExtractedText,
    OfficeConversionQualityReport QualityReport,
    string SummaryNote);

public sealed record DocxHtmlResult(
    string OutputPath,
    string HtmlContent,
    OfficeConversionQualityReport QualityReport,
    string SummaryNote);
