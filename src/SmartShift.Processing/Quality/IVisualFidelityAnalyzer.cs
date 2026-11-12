using SmartShift.Core.Classification;

namespace SmartShift.Processing.Quality;

public interface IVisualFidelityAnalyzer
{
    VisualFidelityMetrics Analyze(
        string sourcePath,
        DetectedFormat sourceFormat,
        string outputPath,
        DetectedFormat outputFormat);
}
