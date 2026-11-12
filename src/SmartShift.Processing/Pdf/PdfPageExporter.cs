using System.Runtime.InteropServices;
using System.Text;
using Docnet.Core;
using Docnet.Core.Exceptions;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using SkiaSharp;
using SmartShift.Core.IO;

namespace SmartShift.Processing.Pdf;

public enum ExportImageFormat
{
    Png,
    Jpeg,
    Webp
}

public sealed record PdfExportResult(
    IReadOnlyList<string> ExportedFiles,
    int TotalPagesInDoc,
    int ExportedCount,
    long TotalBytesWritten,
    ExportImageFormat Format,
    int Dpi,
    string SummaryNote,
    PdfConversionQualityReport? QualityReport = null);

public static class PdfPageExporter
{
    public static PdfExportResult ExportPages(
        string inputPath,
        IOutputPathAllocator allocator,
        ExportImageFormat format,
        int dpi,
        int jpegQuality,
        string? pageRange,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("The source PDF file was not found.", inputPath);
        }

        if (new FileInfo(inputPath).Length == 0)
        {
            throw new InvalidOperationException("The source PDF file is empty (0 bytes).");
        }

        dpi = dpi switch
        {
            72 => 72,
            100 => 100,
            300 => 300,
            _ => 150
        };

        jpegQuality = Math.Clamp(jpegQuality, 1, 100);

        // Standard PDF resolution is 72 points per inch
        double scale = dpi / 72.0;
        var dimensions = new PageDimensions(scale);

        IDocReader docReader;
        try
        {
            docReader = DocLib.Instance.GetDocReader(inputPath, dimensions);
        }
        catch (DocnetLoadDocumentException ex)
        {
            throw new InvalidOperationException(
                "This PDF document is password-protected or uses an unsupported encryption handler. Encrypted PDFs cannot be exported.",
                ex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                "The PDF document is malformed or corrupted and could not be rendered.",
                ex);
        }

        using (docReader)
        {
            int totalPages = docReader.GetPageCount();
            if (totalPages == 0)
            {
                throw new InvalidOperationException("The PDF document contains no pages.");
            }

            var pagesToExport = PageRangeParser.Parse(pageRange, totalPages);
            if (pagesToExport.Count == 0)
            {
                throw new InvalidOperationException("No valid pages were selected for export.");
            }

            allocator.EnsureRoot();
            var baseName = Path.GetFileNameWithoutExtension(inputPath);
            var ext = format switch
            {
                ExportImageFormat.Jpeg => ".jpg",
                ExportImageFormat.Webp => ".webp",
                _ => ".png"
            };
            var exportedFiles = new List<string>();
            var tempPartials = new List<string>();

            try
            {
                int count = pagesToExport.Count;
                for (int i = 0; i < count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    int pageNum = pagesToExport[i];
                    int percent = (int)Math.Round(((i + 1) / (double)count) * 100.0);
                    progress?.Report(new ProcessingProgress(percent, $"Rendering page {pageNum} of {totalPages} ({i + 1}/{count})…"));

                    int pageIndex = pageNum - 1;
                    using var pageReader = docReader.GetPageReader(pageIndex);
                    int width = pageReader.GetPageWidth();
                    int height = pageReader.GetPageHeight();

                    if (ImageSafetyLimits.Exceeds(width, height))
                    {
                        throw new InvalidOperationException(
                            $"Rendered page {pageNum} dimensions ({width} × {height} px) exceed application safety limits.");
                    }

                    byte[] rawBgra = pageReader.GetImage();
                    if (rawBgra == null || rawBgra.Length == 0)
                    {
                        throw new InvalidOperationException($"Failed to render page {pageNum}.");
                    }

                    var desiredName = $"{baseName}-page-{pageNum:D3}{ext}";
                    var finalPath = allocator.Allocate(inputPath, desiredName);
                    var tempPartial = finalPath + ".partial";
                    tempPartials.Add(tempPartial);

                    var info = new SKImageInfo(width, height, SKColorType.Bgra8888, SKAlphaType.Premul);
                    using var bitmap = new SKBitmap(info);
                    Marshal.Copy(rawBgra, 0, bitmap.GetPixels(), rawBgra.Length);

                    if (format == ExportImageFormat.Jpeg)
                    {
                        // Composite over white background so transparency is not converted to black
                        using var surface = SKSurface.Create(info);
                        using var canvas = surface.Canvas;
                        canvas.Clear(SKColors.White);
                        using var skImage = SKImage.FromBitmap(bitmap);
                        canvas.DrawImage(skImage, 0, 0, SKSamplingOptions.Default, null);
                        using var image = surface.Snapshot();
                        using var data = image.Encode(SKEncodedImageFormat.Jpeg, jpegQuality);
                        if (data == null || data.Size == 0)
                        {
                            throw new InvalidOperationException($"Failed to encode page {pageNum} as JPEG.");
                        }
                        File.WriteAllBytes(tempPartial, data.ToArray());
                    }
                    else if (format == ExportImageFormat.Webp)
                    {
                        using var data = bitmap.Encode(SKEncodedImageFormat.Webp, jpegQuality);
                        if (data == null || data.Size == 0)
                        {
                            throw new InvalidOperationException($"Failed to encode page {pageNum} as WebP.");
                        }
                        File.WriteAllBytes(tempPartial, data.ToArray());
                    }
                    else
                    {
                        using var data = bitmap.Encode(SKEncodedImageFormat.Png, 100);
                        if (data == null || data.Size == 0)
                        {
                            throw new InvalidOperationException($"Failed to encode page {pageNum} as PNG.");
                        }
                        File.WriteAllBytes(tempPartial, data.ToArray());
                    }

                    // Validate output image can be decoded and has non-zero dimensions
                    using (var checkCodec = SKCodec.Create(tempPartial))
                    {
                        if (checkCodec == null || checkCodec.Info.Width <= 0 || checkCodec.Info.Height <= 0)
                        {
                            throw new InvalidOperationException($"Output validation failed for rendered page {pageNum}.");
                        }
                    }

                    // Promote atomically
                    File.Move(tempPartial, finalPath, overwrite: false);
                    tempPartials.Remove(tempPartial);
                    exportedFiles.Add(finalPath);
                }

                long totalBytes = exportedFiles.Sum(f => new FileInfo(f).Length);
                string fmtName = format switch
                {
                    ExportImageFormat.Jpeg => "JPEG",
                    ExportImageFormat.Webp => "WebP",
                    _ => "PNG"
                };
                string summary = $"Exported {exportedFiles.Count} of {totalPages} page(s) as {fmtName} ({dpi} DPI).";

                progress?.Report(new ProcessingProgress(100, "Done."));

                var qualityReport = new PdfConversionQualityReport(
                    SourcePath: inputPath,
                    OutputPath: exportedFiles.Count > 0 ? exportedFiles[0] : null,
                    OutputFiles: exportedFiles,
                    ConversionMode: $"PageImageExport_{fmtName}",
                    TotalInputPages: totalPages,
                    PagesProcessed: exportedFiles.Count,
                    TotalCharactersExtracted: 0,
                    ImagesPreservedCount: exportedFiles.Count,
                    TablesReconstructedCount: 0,
                    HeadingsDetectedCount: 0,
                    Status: PdfConversionStatus.Successful,
                    Warnings: [],
                    OutputSizeBytes: totalBytes,
                    SummaryNote: summary);

                return new PdfExportResult(
                    exportedFiles,
                    totalPages,
                    exportedFiles.Count,
                    totalBytes,
                    format,
                    dpi,
                    summary,
                    qualityReport);
            }
            catch
            {
                // Clean up any remaining .partial files
                foreach (var p in tempPartials)
                {
                    TryDelete(p);
                }

                // Clean up any exported files from this failed run to prevent orphaned partial results
                foreach (var f in exportedFiles)
                {
                    TryDelete(f);
                }

                throw;
            }
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
        }
    }
}

public sealed record PdfTextExtractionResult(
    string OutputPath,
    int TotalPagesInDoc,
    int PagesExtractedCount,
    int TotalCharactersExtracted,
    int NonEmptyPagesCount,
    bool IsScannedOrImageOnly,
    string SummaryNote,
    PdfConversionQualityReport? QualityReport = null);

public static class PdfTextExtractor
{
    public static string ExtractAllText(string inputPath)
    {
        if (!File.Exists(inputPath)) return string.Empty;
        try
        {
            using var docReader = DocLib.Instance.GetDocReader(inputPath, new PageDimensions(1.0));
            int totalPages = docReader.GetPageCount();
            var sb = new StringBuilder();
            for (int i = 0; i < totalPages; i++)
            {
                using var pageReader = docReader.GetPageReader(i);
                var text = pageReader.GetText();
                if (!string.IsNullOrWhiteSpace(text))
                {
                    sb.AppendLine(text);
                }
            }
            return sb.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    public static PdfTextExtractionResult ExtractText(
        string inputPath,
        IOutputPathAllocator allocator,
        string? pageRange,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("The source PDF file was not found.", inputPath);
        }

        if (new FileInfo(inputPath).Length == 0)
        {
            throw new InvalidOperationException("The source PDF file is empty (0 bytes).");
        }

        IDocReader docReader;
        try
        {
            docReader = DocLib.Instance.GetDocReader(inputPath, new PageDimensions(1.0));
        }
        catch (DocnetLoadDocumentException ex)
        {
            throw new InvalidOperationException(
                "This PDF document is password-protected or uses an unsupported encryption handler. Encrypted PDFs cannot be processed.",
                ex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                "The PDF document is malformed or corrupted and could not be parsed.",
                ex);
        }

        using (docReader)
        {
            int totalPages = docReader.GetPageCount();
            if (totalPages == 0)
            {
                throw new InvalidOperationException("The PDF document contains no pages.");
            }

            var pagesToExtract = PageRangeParser.Parse(pageRange, totalPages);
            if (pagesToExtract.Count == 0)
            {
                throw new InvalidOperationException("No valid pages were selected for text extraction.");
            }

            allocator.EnsureRoot();
            var desiredName = Path.GetFileNameWithoutExtension(inputPath) + ".txt";
            var finalPath = allocator.Allocate(inputPath, desiredName);
            var tempPartial = finalPath + ".partial";

            try
            {
                var sb = new StringBuilder();
                int totalChars = 0;
                int nonEmptyPages = 0;
                int count = pagesToExtract.Count;

                // Try opening with PdfPig for advanced layout and reading order detection
                UglyToad.PdfPig.PdfDocument? pigDoc = null;
                try
                {
                    pigDoc = UglyToad.PdfPig.PdfDocument.Open(inputPath);
                }
                catch
                {
                    // Fall back to DocLib exclusively if PdfPig cannot open
                }

                using (pigDoc)
                {
                    for (int i = 0; i < count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        int pageNum = pagesToExtract[i];
                        int percent = (int)Math.Round(((i + 1) / (double)count) * 100.0);
                        progress?.Report(new ProcessingProgress(percent, $"Extracting text from page {pageNum} of {totalPages} ({i + 1}/{count})…"));

                        sb.AppendLine($"--- Page {pageNum} ---");
                        string? pageText = null;

                        if (pigDoc != null && pageNum <= pigDoc.NumberOfPages)
                        {
                            try
                            {
                                var pigPage = pigDoc.GetPage(pageNum);
                                var words = pigPage.GetWords().ToList();
                                if (words.Count > 0)
                                {
                                    var blocks = UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter.DocstrumBoundingBoxes.Instance.GetBlocks(words);
                                    if (blocks != null && blocks.Count > 0)
                                    {
                                        var orderedBlocks = UglyToad.PdfPig.DocumentLayoutAnalysis.ReadingOrderDetector.DefaultReadingOrderDetector.Instance.Get(blocks);
                                        var blockStrings = new List<string>();
                                        foreach (var b in orderedBlocks)
                                        {
                                            if (!string.IsNullOrWhiteSpace(b.Text))
                                            {
                                                blockStrings.Add(b.Text.TrimEnd());
                                            }
                                        }
                                        if (blockStrings.Count > 0)
                                        {
                                            pageText = string.Join(Environment.NewLine + Environment.NewLine, blockStrings);
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // Non-fatal layout error, fall back to DocLib text
                            }
                        }

                        // Fallback to DocLib reader if PdfPig yielded nothing
                        if (string.IsNullOrWhiteSpace(pageText))
                        {
                            int pageIndex = pageNum - 1;
                            using var pageReader = docReader.GetPageReader(pageIndex);
                            pageText = pageReader.GetText()?.TrimEnd() ?? "";
                        }

                        if (!string.IsNullOrWhiteSpace(pageText))
                        {
                            sb.AppendLine(pageText);
                            totalChars += pageText.Length;
                            nonEmptyPages++;
                        }
                        else
                        {
                            sb.AppendLine("[No extractable text on this page]");
                        }
                        sb.AppendLine();
                    }
                }

                bool isScanned = (totalChars == 0);

                File.WriteAllText(tempPartial, sb.ToString(), Encoding.UTF8);

                if (!File.Exists(tempPartial) || new FileInfo(tempPartial).Length == 0)
                {
                    throw new InvalidOperationException("Failed to generate text output file.");
                }

                File.Move(tempPartial, finalPath, overwrite: false);

                var warnings = new List<string>();
                var status = PdfConversionStatus.Successful;

                if (isScanned)
                {
                    status = PdfConversionStatus.SuccessfulWithLimitations;
                    warnings.Add("Document appears to be scanned or image-only without an extractable text layer. An OCR engine is required to recognize rasterized text.");
                }

                long outputBytes = new FileInfo(finalPath).Length;
                string summary = isScanned
                    ? $"Extracted 0 characters across {count} page(s). Document appears to be scanned or image-only without an extractable text layer (OCR required)."
                    : $"Extracted {totalChars:N0} characters across {nonEmptyPages} page(s) ({count} page(s) scanned).";

                progress?.Report(new ProcessingProgress(100, "Done."));

                var qualityReport = new PdfConversionQualityReport(
                    SourcePath: inputPath,
                    OutputPath: finalPath,
                    OutputFiles: [finalPath],
                    ConversionMode: "TextExtraction_StructuralOrder",
                    TotalInputPages: totalPages,
                    PagesProcessed: count,
                    TotalCharactersExtracted: totalChars,
                    ImagesPreservedCount: 0,
                    TablesReconstructedCount: 0,
                    HeadingsDetectedCount: 0,
                    Status: status,
                    Warnings: warnings,
                    OutputSizeBytes: outputBytes,
                    SummaryNote: summary);

                return new PdfTextExtractionResult(
                    finalPath,
                    totalPages,
                    count,
                    totalChars,
                    nonEmptyPages,
                    isScanned,
                    summary,
                    qualityReport);
            }
            catch
            {
                TryDelete(tempPartial);
                TryDelete(finalPath);
                throw;
            }
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
        }
    }
}
