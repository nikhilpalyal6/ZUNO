using System.Security.Cryptography;
using PdfSharp.Pdf.IO;
using SmartShift.Core.IO;
using SmartShift.Processing;
using SmartShift.Processing.Documents.Editing;

namespace SmartShift.Processing.Pdf.Editing;

public sealed class PdfMetadataService : IPdfMetadataService
{
    private readonly IOutputPathAllocator _allocator;

    public PdfMetadataService(IOutputPathAllocator allocator)
    {
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
    }

    public async Task<PdfMetadata> ReadMetadataAsync(string inputPath, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
        {
            throw new FileNotFoundException($"Input PDF file not found: {inputPath}", inputPath);
        }

        cancellationToken.ThrowIfCancellationRequested();

        return await Task.Run(() =>
        {
            using var doc = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
            return new PdfMetadata(
                Title: doc.Info.Title,
                Author: doc.Info.Author,
                Subject: doc.Info.Subject,
                Keywords: doc.Info.Keywords,
                Creator: doc.Info.Creator,
                Producer: doc.Info.Producer);
        }, cancellationToken).ConfigureAwait(false);
    }

    public async Task<DocumentEditingResult> UpdateMetadataAsync(
        string inputPath,
        string outputPath,
        PdfMetadata metadata,
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

        ArgumentNullException.ThrowIfNull(metadata);
        cancellationToken.ThrowIfCancellationRequested();

        return await Task.Run(() => UpdateMetadataInternal(inputPath, outputPath, metadata, progress, cancellationToken), cancellationToken)
            .ConfigureAwait(false);
    }

    private static DocumentEditingResult UpdateMetadataInternal(
        string inputPath,
        string outputPath,
        PdfMetadata metadata,
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
            progress?.Report(new ProcessingProgress(30, "Reading document information dictionary…"));
            File.Copy(inputPath, tempPartial, overwrite: true);

            using (var doc = PdfReader.Open(tempPartial, PdfDocumentOpenMode.Modify))
            {
                if (doc.SecuritySettings.IsEncrypted)
                {
                    throw new InvalidOperationException("Encrypted or password-protected PDF files cannot be edited.");
                }

                progress?.Report(new ProcessingProgress(60, "Writing updated metadata fields…"));

                if (metadata.Title != null) doc.Info.Title = metadata.Title;
                if (metadata.Author != null) doc.Info.Author = metadata.Author;
                if (metadata.Subject != null) doc.Info.Subject = metadata.Subject;
                if (metadata.Keywords != null) doc.Info.Keywords = metadata.Keywords;
                if (metadata.Creator != null) doc.Info.Creator = metadata.Creator;
                if (metadata.Producer != null) doc.Info.Elements.SetString("/Producer", metadata.Producer);

                doc.Save(tempPartial);
            }

            progress?.Report(new ProcessingProgress(90, "Finalizing updated PDF…"));

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
            File.Move(tempPartial, outputPath, overwrite: true);

            var postHash = ComputeSha256(inputPath);
            if (!string.Equals(preHash, postHash, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"CRITICAL: Source file '{inputPath}' was modified during metadata operation!");
            }

            progress?.Report(new ProcessingProgress(100, "PDF metadata updated successfully."));

            return new DocumentEditingResult(
                OutputPath: outputPath,
                IsSuccess: true,
                Limitations: [],
                Message: "PDF metadata dictionary successfully updated.");
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

    private static string ComputeSha256(string path)
    {
        using var sha = SHA256.Create();
        using var stream = File.OpenRead(path);
        var bytes = sha.ComputeHash(stream);
        return Convert.ToHexString(bytes);
    }
}
