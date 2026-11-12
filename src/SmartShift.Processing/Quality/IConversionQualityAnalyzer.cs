using SmartShift.Core.Classification;

namespace SmartShift.Processing.Quality;

public interface IConversionQualityAnalyzer
{
    Task<ConversionQualityReport> AnalyzeAsync(
        string sourcePath,
        DetectedFormat sourceFormat,
        string outputPath,
        DetectedFormat outputFormat,
        string operationId,
        ValidationLevel level = ValidationLevel.Standard,
        CancellationToken cancellationToken = default);
}
