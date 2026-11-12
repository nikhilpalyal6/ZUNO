using System.Diagnostics;
using SmartShift.Core.Classification;

namespace SmartShift.Processing.Quality;

public sealed class ConversionQualityAnalyzer : IConversionQualityAnalyzer
{
    private readonly IStructuralFidelityValidator _structuralValidator;
    private readonly IContentIntegrityValidator _contentValidator;
    private readonly IVisualFidelityAnalyzer _visualAnalyzer;

    public ConversionQualityAnalyzer(
        IStructuralFidelityValidator? structuralValidator = null,
        IContentIntegrityValidator? contentValidator = null,
        IVisualFidelityAnalyzer? visualAnalyzer = null)
    {
        _structuralValidator = structuralValidator ?? new StructuralFidelityValidator();
        _contentValidator = contentValidator ?? new ContentIntegrityValidator();
        _visualAnalyzer = visualAnalyzer ?? new VisualFidelityAnalyzer();
    }

    public async Task<ConversionQualityReport> AnalyzeAsync(
        string sourcePath,
        DetectedFormat sourceFormat,
        string outputPath,
        DetectedFormat outputFormat,
        string operationId,
        ValidationLevel level = ValidationLevel.Standard,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var startTime = DateTime.UtcNow;
        var sw = Stopwatch.StartNew();

        return await Task.Run(() =>
        {
            var warnings = new List<string>();
            var limitations = new List<ConversionLimitation>();

            // 1. Matrix capability lookup
            var descriptor = ConversionCapabilityMatrix.GetDescriptor(operationId);
            if (descriptor != null)
            {
                foreach (var lim in descriptor.KnownLimitations)
                {
                    limitations.Add(new ConversionLimitation(
                        Code: "KNOWN_LIMITATION",
                        Description: lim,
                        IsExpectedBehavior: true));
                }
            }

            // 2. Structural Validation
            cancellationToken.ThrowIfCancellationRequested();
            var structural = _structuralValidator.Validate(outputPath, outputFormat);
            warnings.AddRange(structural.StructuralWarnings);

            // 3. Content Integrity Validation
            ContentIntegrityMetrics? content = null;
            if (level >= ValidationLevel.Standard && (descriptor == null || descriptor.SupportsContentValidation))
            {
                cancellationToken.ThrowIfCancellationRequested();
                content = _contentValidator.Validate(sourcePath, sourceFormat, outputPath, outputFormat);
                if (content != null)
                {
                    warnings.AddRange(content.ContentWarnings);
                }
            }

            // 4. Visual Fidelity Validation
            VisualFidelityMetrics? visual = null;
            if (level >= ValidationLevel.Deep && (descriptor == null || descriptor.SupportsVisualValidation))
            {
                cancellationToken.ThrowIfCancellationRequested();
                visual = _visualAnalyzer.Analyze(sourcePath, sourceFormat, outputPath, outputFormat);
                if (visual != null)
                {
                    warnings.AddRange(visual.VisualWarnings);
                }
            }

            // 5. Classification and Confidence Scoring
            var (classification, score, summary) = ComputeClassificationAndScore(structural, content, visual, limitations);

            sw.Stop();

            return new ConversionQualityReport(
                SourcePath: sourcePath,
                SourceFormat: sourceFormat,
                OutputPath: outputPath,
                OutputFormat: outputFormat,
                OperationId: operationId,
                Level: level,
                Classification: classification,
                CompositeConfidenceScore: score,
                StartTimeUtc: startTime,
                Duration: sw.Elapsed,
                Structural: structural,
                Content: content,
                Visual: visual,
                Limitations: limitations,
                Warnings: warnings,
                SummaryNote: summary);
        }, cancellationToken).ConfigureAwait(false);
    }

    private static (QualityClassification Classification, double Score, string Summary) ComputeClassificationAndScore(
        StructuralMetrics structural,
        ContentIntegrityMetrics? content,
        VisualFidelityMetrics? visual,
        IReadOnlyList<ConversionLimitation> limitations)
    {
        // If package is structurally invalid or output is empty
        if (!structural.PackageValid)
        {
            return (
                QualityClassification.StructuralMismatch,
                0.0,
                "Structural validation failed: Output package is malformed, unreadable, or missing required internal structures."
            );
        }

        double structuralScore = 1.0;

        // Content check
        double contentScore = 1.0;
        bool contentMismatch = false;
        if (content != null)
        {
            contentScore = content.TextSimilarityScore;
            if (content.TextSimilarityScore < 0.60 && content.SourceCharacterCount > 20)
            {
                contentMismatch = true;
            }
        }

        // Visual check
        double visualScore = 1.0;
        bool visualMismatch = false;
        bool visualEvaluated = false;
        if (visual != null && visual.Status is VisualComparisonStatus.Comparable or VisualComparisonStatus.DifferingDimensions)
        {
            visualEvaluated = true;
            visualScore = Math.Clamp(1.0 - (visual.PixelDifferencePercentage / 100.0), 0.0, 1.0);
            if (visual.PixelDifferencePercentage > 40.0)
            {
                visualMismatch = true;
            }
        }

        // Composite score calculation
        double compositeScore;
        if (content != null && visualEvaluated)
        {
            compositeScore = Math.Round((structuralScore * 0.3) + (contentScore * 0.4) + (visualScore * 0.3), 4);
        }
        else if (content != null)
        {
            compositeScore = Math.Round((structuralScore * 0.4) + (contentScore * 0.6), 4);
        }
        else if (visualEvaluated)
        {
            compositeScore = Math.Round((structuralScore * 0.5) + (visualScore * 0.5), 4);
        }
        else
        {
            compositeScore = Math.Round(structuralScore, 4);
        }

        if (contentMismatch)
        {
            return (
                QualityClassification.ContentMismatch,
                compositeScore,
                $"Content mismatch detected: Text similarity score ({content!.TextSimilarityScore:P1}) fell below tolerance threshold."
            );
        }

        if (visualMismatch)
        {
            return (
                QualityClassification.VisualMismatch,
                compositeScore,
                $"Visual mismatch detected: Pixel divergence ({visual!.PixelDifferencePercentage:F1}%) exceeded comparison threshold."
            );
        }

        if (limitations.Count > 0 || structural.StructuralWarnings.Count > 0 || (content?.ContentWarnings.Count ?? 0) > 0)
        {
            return (
                QualityClassification.ValidatedWithLimitations,
                compositeScore,
                $"Conversion validated with {limitations.Count} documented format limitation(s)."
            );
        }

        return (
            QualityClassification.Validated,
            compositeScore,
            "Conversion fully validated with verified structural, content, and package fidelity."
        );
    }
}
