using SmartShift.Core.Classification;

namespace SmartShift.Processing.Quality;

public interface IContentIntegrityValidator
{
    ContentIntegrityMetrics? Validate(
        string sourcePath,
        DetectedFormat sourceFormat,
        string outputPath,
        DetectedFormat outputFormat);
}
