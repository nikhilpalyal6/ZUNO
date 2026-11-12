using System.Runtime.InteropServices;
using SkiaSharp;
using SmartShift.Core.Classification;
using SmartShift.Processing.Pdf.Preview;

namespace SmartShift.Processing.Quality;

public sealed class VisualFidelityAnalyzer : IVisualFidelityAnalyzer
{
    private const double ToleranceThreshold = 18.0; // Perceptual RGB Euclidean distance threshold

    public VisualFidelityMetrics Analyze(
        string sourcePath,
        DetectedFormat sourceFormat,
        string outputPath,
        DetectedFormat outputFormat)
    {
        if (!IsFormatVisuallyRenderable(sourceFormat) || !IsFormatVisuallyRenderable(outputFormat))
        {
            var nonRenderable = !IsFormatVisuallyRenderable(sourceFormat) ? sourceFormat : outputFormat;
            return new VisualFidelityMetrics(
                Status: VisualComparisonStatus.NotComparable,
                ComparisonMethod: "None",
                ComparedPagesCount: 0,
                MeanAbsoluteError: 0.0,
                RootMeanSquareError: 0.0,
                PixelDifferencePercentage: 0.0,
                Rationale: $"Visual raster comparison is not supported for {nonRenderable}. Layout flow is determined dynamically by document viewer engines.",
                VisualWarnings: []);
        }

        if (!File.Exists(sourcePath) || !File.Exists(outputPath))
        {
            return new VisualFidelityMetrics(
                Status: VisualComparisonStatus.NotComparable,
                ComparisonMethod: "None",
                ComparedPagesCount: 0,
                MeanAbsoluteError: 0.0,
                RootMeanSquareError: 0.0,
                PixelDifferencePercentage: 0.0,
                Rationale: "Source or output file does not exist.",
                VisualWarnings: ["Missing file during visual fidelity analysis."]);
        }

        try
        {
            using var sourceBitmap = RenderOrLoadBitmap(sourcePath, sourceFormat);
            using var targetBitmap = RenderOrLoadBitmap(outputPath, outputFormat);

            if (sourceBitmap == null || targetBitmap == null)
            {
                return new VisualFidelityMetrics(
                    Status: VisualComparisonStatus.NotComparable,
                    ComparisonMethod: "None",
                    ComparedPagesCount: 0,
                    MeanAbsoluteError: 0.0,
                    RootMeanSquareError: 0.0,
                    PixelDifferencePercentage: 0.0,
                    Rationale: "Failed to decode raster bitmaps for visual comparison.",
                    VisualWarnings: ["Bitmap decode or page render returned null."]);
            }

            return CompareBitmaps(sourceBitmap, targetBitmap);
        }
        catch (Exception ex)
        {
            return new VisualFidelityMetrics(
                Status: VisualComparisonStatus.NotComparable,
                ComparisonMethod: "Exception",
                ComparedPagesCount: 0,
                MeanAbsoluteError: 0.0,
                RootMeanSquareError: 0.0,
                PixelDifferencePercentage: 0.0,
                Rationale: $"Visual comparison exception: {ex.Message}",
                VisualWarnings: [ex.Message]);
        }
    }

    private static bool IsFormatVisuallyRenderable(DetectedFormat format) =>
        format is DetectedFormat.Pdf or DetectedFormat.Png or DetectedFormat.Jpeg or DetectedFormat.Webp;

    private static SKBitmap? RenderOrLoadBitmap(string path, DetectedFormat format)
    {
        if (format is DetectedFormat.Png or DetectedFormat.Jpeg or DetectedFormat.Webp)
        {
            return SKBitmap.Decode(path);
        }

        if (format == DetectedFormat.Pdf)
        {
            // Render first page at 96 DPI for bounded memory and quick comparison
            var task = PdfPreviewRenderer.RenderPageAsync(path, 0, 96.0);
            task.Wait();
            var pageData = task.Result;

            var bitmap = new SKBitmap(pageData.PixelWidth, pageData.PixelHeight, SKColorType.Bgra8888, SKAlphaType.Premul);
            Marshal.Copy(pageData.BgraBytes, 0, bitmap.GetPixels(), pageData.BgraBytes.Length);
            return bitmap;
        }

        return null;
    }

    public static VisualFidelityMetrics CompareBitmaps(SKBitmap source, SKBitmap target)
    {
        var warnings = new List<string>();
        var status = VisualComparisonStatus.Comparable;

        SKBitmap? resizedTarget = null;
        SKBitmap evalTarget = target;

        if (source.Width != target.Width || source.Height != target.Height)
        {
            status = VisualComparisonStatus.DifferingDimensions;
            warnings.Add($"Dimension discrepancy: Source is {source.Width}x{source.Height}, Output is {target.Width}x{target.Height}. Resampling output for pixel analysis.");

            resizedTarget = target.Resize(new SKImageInfo(source.Width, source.Height), new SKSamplingOptions(SKCubicResampler.Mitchell));
            evalTarget = resizedTarget ?? target;
        }

        try
        {
            int totalPixels = source.Width * source.Height;
            if (totalPixels == 0)
            {
                return new VisualFidelityMetrics(
                    Status: VisualComparisonStatus.NotComparable,
                    ComparisonMethod: "ZeroPixels",
                    ComparedPagesCount: 0,
                    MeanAbsoluteError: 0.0,
                    RootMeanSquareError: 0.0,
                    PixelDifferencePercentage: 0.0,
                    Rationale: "Source bitmap has 0 pixels.",
                    VisualWarnings: warnings);
            }

            double totalAbsDiff = 0.0;
            double totalSqDiff = 0.0;
            int differingPixels = 0;

            for (int y = 0; y < source.Height; y++)
            {
                for (int x = 0; x < source.Width; x++)
                {
                    var c1 = source.GetPixel(x, y);
                    var c2 = evalTarget.GetPixel(x, y);

                    double dr = c1.Red - c2.Red;
                    double dg = c1.Green - c2.Green;
                    double db = c1.Blue - c2.Blue;
                    double da = c1.Alpha - c2.Alpha;

                    double absDiff = (Math.Abs(dr) + Math.Abs(dg) + Math.Abs(db) + Math.Abs(da)) / (4.0 * 255.0);
                    double sqDiff = (dr * dr + dg * dg + db * db + da * da) / (4.0 * 255.0 * 255.0);

                    totalAbsDiff += absDiff;
                    totalSqDiff += sqDiff;

                    double euclideanDistance = Math.Sqrt(dr * dr + dg * dg + db * db);
                    if (euclideanDistance > ToleranceThreshold)
                    {
                        differingPixels++;
                    }
                }
            }

            double mae = Math.Round(totalAbsDiff / totalPixels, 4);
            double rmse = Math.Round(Math.Sqrt(totalSqDiff / totalPixels), 4);
            double diffPct = Math.Round((double)differingPixels / totalPixels * 100.0, 2);

            return new VisualFidelityMetrics(
                Status: status,
                ComparisonMethod: "NormalizedRasterDiff(MAE/RMSE/Euclidean)",
                ComparedPagesCount: 1,
                MeanAbsoluteError: mae,
                RootMeanSquareError: rmse,
                PixelDifferencePercentage: diffPct,
                Rationale: null,
                VisualWarnings: warnings);
        }
        finally
        {
            resizedTarget?.Dispose();
        }
    }
}
