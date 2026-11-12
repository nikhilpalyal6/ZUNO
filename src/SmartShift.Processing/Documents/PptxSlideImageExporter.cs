using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Docnet.Core;
using Docnet.Core.Models;
using SkiaSharp;
using SmartShift.Core.IO;
using SmartShift.Processing;

namespace SmartShift.Processing.Documents;

public static class PptxSlideImageExporter
{
    public static OfficeMultiOutputResult Export(
        string sourcePptxPath,
        SlideExportFormat format = SlideExportFormat.Png,
        int dpi = 150,
        int quality = 85,
        string slideRange = "all",
        string? outputDirectory = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (string.IsNullOrWhiteSpace(sourcePptxPath) || !File.Exists(sourcePptxPath))
        {
            throw new FileNotFoundException("Source PowerPoint presentation does not exist.", sourcePptxPath);
        }

        var sw = Stopwatch.StartNew();
        outputDirectory ??= Path.GetDirectoryName(sourcePptxPath) ?? Directory.GetCurrentDirectory();
        Directory.CreateDirectory(outputDirectory);

        var stagingDir = Path.Combine(outputDirectory, $".staging_pptx_img_{Guid.NewGuid():N}");
        Directory.CreateDirectory(stagingDir);

        var tempPdfPath = Path.Combine(stagingDir, "intermediate_deck.pdf");
        var generatedFiles = new List<string>();
        var limitations = new List<string>();
        var warnings = new List<string>();

        try
        {
            progress?.Report(new ProcessingProgress(10, "Converting presentation slides to vector layout…"));
            // Step 1: Render slides to intermediate PDF
            var pdfResult = PptxToPdfConverter.Convert(sourcePptxPath, tempPdfPath, slideRange, progress, cancellationToken);
            limitations.AddRange(pdfResult.QualityReport.Limitations);
            warnings.AddRange(pdfResult.QualityReport.Warnings);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ProcessingProgress(50, "Rasterizing presentation slides at target resolution…"));

            var ext = format switch
            {
                SlideExportFormat.Jpeg => ".jpg",
                SlideExportFormat.Webp => ".webp",
                _ => ".png"
            };

            var skFormat = format switch
            {
                SlideExportFormat.Jpeg => SKEncodedImageFormat.Jpeg,
                SlideExportFormat.Webp => SKEncodedImageFormat.Webp,
                _ => SKEncodedImageFormat.Png
            };

            var baseFileName = Path.GetFileNameWithoutExtension(sourcePptxPath);
            double scale = dpi / 72.0;
            var dimensions = new PageDimensions(scale);

            using (var docReader = DocLib.Instance.GetDocReader(tempPdfPath, dimensions))
            {
                int pageCount = docReader.GetPageCount();
                if (pageCount == 0)
                {
                    throw new InvalidOperationException("Rendered presentation contains no pages.");
                }

                for (int p = 0; p < pageCount; p++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int pct = 50 + (int)(40.0 * (p + 1) / pageCount);
                    progress?.Report(new ProcessingProgress(pct, $"Exporting slide image {p + 1} of {pageCount} ({format})…"));

                    using var pageReader = docReader.GetPageReader(p);
                    int targetWidth = pageReader.GetPageWidth();
                    int targetHeight = pageReader.GetPageHeight();

                    var rawBytes = pageReader.GetImage();
                    if (rawBytes == null || rawBytes.Length == 0)
                    {
                        throw new InvalidOperationException($"Failed to render slide page {p + 1}.");
                    }

                    var info = new SKImageInfo(targetWidth, targetHeight, SKColorType.Bgra8888, SKAlphaType.Premul);
                    using var bitmap = new SKBitmap(info);
                    Marshal.Copy(rawBytes, 0, bitmap.GetPixels(), rawBytes.Length);

                    string slideFileName = $"{baseFileName}_slide_{p + 1:D2}{ext}";
                    string stagedFilePath = Path.Combine(stagingDir, slideFileName);

                    if (format == SlideExportFormat.Jpeg)
                    {
                        // Composite over white background so transparency is not converted to black
                        using var surface = SKSurface.Create(info);
                        using var canvas = surface.Canvas;
                        canvas.Clear(SKColors.White);
                        using var skImg = SKImage.FromBitmap(bitmap);
                        canvas.DrawImage(skImg, 0, 0, SKSamplingOptions.Default, null);
                        using var image = surface.Snapshot();
                        using var data = image.Encode(skFormat, Math.Clamp(quality, 1, 100));

                        using var fs = new FileStream(stagedFilePath, FileMode.Create, FileAccess.Write, FileShare.None);
                        data.SaveTo(fs);
                    }
                    else
                    {
                        using var skImg = SKImage.FromBitmap(bitmap);
                        using var data = skImg.Encode(skFormat, Math.Clamp(quality, 1, 100));

                        using var fs = new FileStream(stagedFilePath, FileMode.Create, FileAccess.Write, FileShare.None);
                        data.SaveTo(fs);
                    }

                    // Validate generated image
                    using (var codec = SKCodec.Create(stagedFilePath))
                    {
                        if (codec == null || codec.Info.Width <= 0 || codec.Info.Height <= 0)
                        {
                            throw new InvalidOperationException($"Exported image '{slideFileName}' is corrupt or unreadable.");
                        }
                    }

                    // Promote from staging to final output path
                    string finalPath = Path.Combine(outputDirectory, slideFileName);
                    if (File.Exists(finalPath))
                    {
                        File.Delete(finalPath);
                    }
                    File.Move(stagedFilePath, finalPath);
                    generatedFiles.Add(finalPath);
                }
            }

            sw.Stop();
            progress?.Report(new ProcessingProgress(100, $"Successfully exported {generatedFiles.Count} slide image(s)."));

            var status = limitations.Count > 0 ? OfficeConversionStatus.SuccessfulWithLimitations : OfficeConversionStatus.Successful;
            var report = new OfficeConversionQualityReport(
                Status: status,
                ConversionMode: "PptxSlideImageExport",
                PagesOrSlidesCount: generatedFiles.Count,
                HeadingsDetectedCount: pdfResult.QualityReport.HeadingsDetectedCount,
                TablesReconstructedCount: pdfResult.QualityReport.TablesReconstructedCount,
                ImagesPreservedCount: generatedFiles.Count,
                ShapesPreservedCount: pdfResult.QualityReport.ShapesPreservedCount,
                CharacterCount: pdfResult.QualityReport.CharacterCount,
                Limitations: limitations,
                Warnings: warnings,
                ValidationPassed: true,
                Duration: sw.Elapsed);

            return new OfficeMultiOutputResult(
                generatedFiles,
                report,
                $"Exported {generatedFiles.Count} slide(s) to {format} at {dpi} DPI.");
        }
        catch
        {
            // Clean up any partially promoted outputs on error or cancellation
            foreach (var file in generatedFiles)
            {
                try { if (File.Exists(file)) File.Delete(file); } catch { }
            }
            throw;
        }
        finally
        {
            // Always purge staging directory
            try
            {
                if (Directory.Exists(stagingDir))
                {
                    Directory.Delete(stagingDir, true);
                }
            }
            catch { }
        }
    }
}
