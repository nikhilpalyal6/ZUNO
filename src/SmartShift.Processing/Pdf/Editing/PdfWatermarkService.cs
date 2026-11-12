using System.Globalization;
using System.Security.Cryptography;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using SmartShift.Core.IO;
using SmartShift.Processing;
using SmartShift.Processing.Documents.Editing;

namespace SmartShift.Processing.Pdf.Editing;

public sealed class PdfWatermarkService : IPdfWatermarkService
{
    private readonly IOutputPathAllocator _allocator;

    public PdfWatermarkService(IOutputPathAllocator allocator)
    {
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
    }

    public async Task<DocumentEditingResult> AddWatermarkAsync(
        string inputPath,
        string outputPath,
        PdfWatermarkOptions options,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
        {
            throw new FileNotFoundException($"Input PDF file not found: {inputPath}", inputPath);
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            throw new ArgumentException("Output path must be specified.", nameof(outputPath));
        }

        ArgumentNullException.ThrowIfNull(options);
        if (string.IsNullOrWhiteSpace(options.Text))
        {
            throw new ArgumentException("Watermark text cannot be empty.", nameof(options));
        }

        cancellationToken.ThrowIfCancellationRequested();

        return await Task.Run(() => AddWatermarkInternal(inputPath, outputPath, options, progress, cancellationToken), cancellationToken)
            .ConfigureAwait(false);
    }

    private static DocumentEditingResult AddWatermarkInternal(
        string inputPath,
        string outputPath,
        PdfWatermarkOptions options,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        progress?.Report(new ProcessingProgress(10, "Verifying PDF document…"));
        var preHash = ComputeSha256(inputPath);

        var dir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        var tempPartial = outputPath + ".partial";
        if (File.Exists(tempPartial))
        {
            try { File.Delete(tempPartial); } catch { }
        }

        try
        {
            progress?.Report(new ProcessingProgress(25, "Loading PDF structure…"));
            File.Copy(inputPath, tempPartial, overwrite: true);

            using (var doc = PdfReader.Open(tempPartial, PdfDocumentOpenMode.Modify))
            {
                if (doc.SecuritySettings.IsEncrypted)
                {
                    throw new InvalidOperationException("Encrypted or password-protected PDF files cannot be watermarked.");
                }

                if (doc.PageCount == 0)
                {
                    throw new InvalidOperationException("The PDF document contains 0 pages.");
                }

                progress?.Report(new ProcessingProgress(50, "Rendering watermark on requested pages…"));

                var targetPageIndices = ResolveTargetPages(doc.PageCount, options.PageRange);
                var (r, g, b) = ParseColor(options.HexColor);
                var alpha = (int)Math.Round(Math.Clamp(options.Opacity, 0.05, 1.0) * 255.0);
                var brush = new XSolidBrush(XColor.FromArgb(alpha, r, g, b));
                var fontName = string.IsNullOrWhiteSpace(options.FontName) || options.FontName.Equals("Helvetica", StringComparison.OrdinalIgnoreCase)
                    ? "Arial"
                    : options.FontName;
                XFont font;
                try
                {
                    font = new XFont(fontName, options.FontSize, XFontStyleEx.Bold);
                }
                catch
                {
                    font = new XFont("Arial", options.FontSize, XFontStyleEx.Bold);
                }
                var format = new XStringFormat
                {
                    Alignment = XStringAlignment.Center,
                    LineAlignment = XLineAlignment.Center
                };

                for (int i = 0; i < doc.PageCount; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (!targetPageIndices.Contains(i))
                    {
                        continue;
                    }

                    var page = doc.Pages[i];
                    using var gfx = XGraphics.FromPdfPage(page);

                    double cx = page.Width.Point / 2.0;
                    double cy = page.Height.Point / 2.0;
                    double angle = options.IsDiagonal
                        ? -Math.Atan2(page.Height.Point, page.Width.Point) * (180.0 / Math.PI)
                        : 0.0;

                    var state = gfx.Save();
                    gfx.TranslateTransform(cx, cy);
                    gfx.RotateTransform(angle);
                    gfx.DrawString(options.Text, font, brush, 0, 0, format);
                    gfx.Restore(state);
                }

                doc.Save(tempPartial);
            }

            progress?.Report(new ProcessingProgress(90, "Finalizing watermarked PDF…"));

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
            File.Move(tempPartial, outputPath, overwrite: true);

            var postHash = ComputeSha256(inputPath);
            if (!string.Equals(preHash, postHash, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"CRITICAL: Source file '{inputPath}' was modified during watermark operation!");
            }

            progress?.Report(new ProcessingProgress(100, "PDF watermark completed successfully."));

            return new DocumentEditingResult(
                OutputPath: outputPath,
                IsSuccess: true,
                Limitations: [],
                Message: $"Watermark '{options.Text}' successfully applied to PDF.");
        }
        catch
        {
            if (File.Exists(tempPartial))
            {
                try { File.Delete(tempPartial); } catch { }
            }
            throw;
        }
    }

    private static HashSet<int> ResolveTargetPages(int totalPages, string? pageRange)
    {
        var result = new HashSet<int>();
        if (string.IsNullOrWhiteSpace(pageRange) || pageRange.Trim().Equals("all", StringComparison.OrdinalIgnoreCase))
        {
            for (int i = 0; i < totalPages; i++) result.Add(i);
            return result;
        }

        var trimmed = pageRange.Trim().ToLowerInvariant();
        if (trimmed == "odd")
        {
            for (int i = 0; i < totalPages; i += 2) result.Add(i);
            return result;
        }

        if (trimmed == "even")
        {
            for (int i = 1; i < totalPages; i += 2) result.Add(i);
            return result;
        }

        var parts = trimmed.Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var part in parts)
        {
            if (part.Contains('-'))
            {
                var dashParts = part.Split('-', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                if (dashParts.Length == 2 &&
                    int.TryParse(dashParts[0], out var start) &&
                    int.TryParse(dashParts[1], out var end))
                {
                    int min = Math.Max(1, Math.Min(start, end));
                    int max = Math.Min(totalPages, Math.Max(start, end));
                    for (int p = min; p <= max; p++)
                    {
                        result.Add(p - 1);
                    }
                }
            }
            else if (int.TryParse(part, out var singlePage) && singlePage >= 1 && singlePage <= totalPages)
            {
                result.Add(singlePage - 1);
            }
        }

        if (result.Count == 0)
        {
            for (int i = 0; i < totalPages; i++) result.Add(i);
        }

        return result;
    }

    private static (byte R, byte G, byte B) ParseColor(string? hex)
    {
        if (string.IsNullOrWhiteSpace(hex)) return (128, 128, 128);
        var clean = hex.Trim().TrimStart('#');

        if (clean.Length == 6 &&
            byte.TryParse(clean[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var r) &&
            byte.TryParse(clean[2..4], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var g) &&
            byte.TryParse(clean[4..6], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
        {
            return (r, g, b);
        }

        return clean.ToLowerInvariant() switch
        {
            "red" => (255, 0, 0),
            "green" => (0, 128, 0),
            "blue" => (0, 0, 255),
            "black" => (0, 0, 0),
            _ => (128, 128, 128)
        };
    }

    private static string ComputeSha256(string path)
    {
        using var sha = SHA256.Create();
        using var stream = File.OpenRead(path);
        var bytes = sha.ComputeHash(stream);
        return Convert.ToHexString(bytes);
    }
}
