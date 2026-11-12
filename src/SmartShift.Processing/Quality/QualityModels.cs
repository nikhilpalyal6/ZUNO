using System.Text.Json.Serialization;
using SmartShift.Core.Classification;

namespace SmartShift.Processing.Quality;

public enum ValidationLevel
{
    /// <summary>
    /// Package integrity, file readability, and schema/format structure.
    /// </summary>
    Basic = 0,

    /// <summary>
    /// Structural checks plus text/table content extraction and similarity comparison.
    /// </summary>
    Standard = 1,

    /// <summary>
    /// Structural, content, and rasterized visual page comparison (where technically supported).
    /// </summary>
    Deep = 2
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum QualityClassification
{
    Validated,
    ValidatedWithLimitations,
    ContentMismatch,
    StructuralMismatch,
    VisualMismatch,
    ValidationIncomplete,
    ValidationFailed,
    NotComparable
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum VisualComparisonStatus
{
    Comparable,
    NotComparable,
    DifferingDimensions,
    Sampled
}

public sealed record StructuralMetrics(
    bool PackageValid,
    int ElementCount,
    int PageOrSlideCount,
    int ParagraphCount,
    int HeadingCount,
    int TableCount,
    int ImageCount,
    int ShapeCount,
    int SectionCount,
    IReadOnlyList<string> StructuralWarnings);

public sealed record ContentIntegrityMetrics(
    long SourceCharacterCount,
    long OutputCharacterCount,
    double TextSimilarityScore,
    int MissingTextSegments,
    int DuplicateTextSegments,
    bool UnicodePreserved,
    bool TablesPreserved,
    IReadOnlyList<string> ContentWarnings);

public sealed record VisualFidelityMetrics(
    VisualComparisonStatus Status,
    string ComparisonMethod,
    int ComparedPagesCount,
    double MeanAbsoluteError,
    double RootMeanSquareError,
    double PixelDifferencePercentage,
    string? Rationale,
    IReadOnlyList<string> VisualWarnings);

public sealed record ConversionLimitation(
    string Code,
    string Description,
    bool IsExpectedBehavior);

public sealed record ConversionQualityReport(
    string SourcePath,
    DetectedFormat SourceFormat,
    string OutputPath,
    DetectedFormat OutputFormat,
    string OperationId,
    ValidationLevel Level,
    QualityClassification Classification,
    double CompositeConfidenceScore,
    DateTime StartTimeUtc,
    TimeSpan Duration,
    StructuralMetrics Structural,
    ContentIntegrityMetrics? Content,
    VisualFidelityMetrics? Visual,
    IReadOnlyList<ConversionLimitation> Limitations,
    IReadOnlyList<string> Warnings,
    string SummaryNote)
{
    public bool IsAcceptable => Classification is QualityClassification.Validated or QualityClassification.ValidatedWithLimitations;
}
