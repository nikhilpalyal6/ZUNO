using SmartShift.Core.Classification;

namespace SmartShift.Processing.Quality;

public interface IStructuralFidelityValidator
{
    StructuralMetrics Validate(string filePath, DetectedFormat format);
}
