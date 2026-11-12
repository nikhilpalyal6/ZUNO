using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using SmartShift.Core.IO;
using SmartShift.Processing;

namespace SmartShift.Processing.Documents.Editing;

public sealed class DocxMetadataService : IDocxMetadataService
{
    private readonly IOutputPathAllocator _allocator;

    public DocxMetadataService(IOutputPathAllocator allocator)
    {
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
    }

    public async Task<DocxMetadata> ReadMetadataAsync(string inputPath, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
        {
            throw new FileNotFoundException($"Input DOCX file not found: {inputPath}", inputPath);
        }

        cancellationToken.ThrowIfCancellationRequested();

        return await Task.Run(() =>
        {
            using var doc = WordprocessingDocument.Open(inputPath, false);
            var props = doc.PackageProperties;
            return new DocxMetadata(
                Title: props.Title,
                Subject: props.Subject,
                Author: props.Creator,
                Keywords: props.Keywords,
                Comments: props.Description,
                Category: props.Category);
        }, cancellationToken).ConfigureAwait(false);
    }

    public async Task<DocumentEditingResult> UpdateMetadataAsync(
        string inputPath,
        string outputPath,
        DocxMetadata metadata,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
        {
            throw new FileNotFoundException($"Input DOCX file not found: {inputPath}", inputPath);
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
        DocxMetadata metadata,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        progress?.Report(new ProcessingProgress(10, "Verifying input document…"));
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
            progress?.Report(new ProcessingProgress(30, "Applying metadata changes…"));
            File.Copy(inputPath, tempPartial, overwrite: true);

            using (var doc = WordprocessingDocument.Open(tempPartial, true))
            {
                var props = doc.PackageProperties;
                if (metadata.Title != null) props.Title = metadata.Title;
                if (metadata.Subject != null) props.Subject = metadata.Subject;
                if (metadata.Author != null) props.Creator = metadata.Author;
                if (metadata.Keywords != null) props.Keywords = metadata.Keywords;
                if (metadata.Comments != null) props.Description = metadata.Comments;
                if (metadata.Category != null) props.Category = metadata.Category;
            }

            progress?.Report(new ProcessingProgress(85, "Finalizing updated document…"));

            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
            File.Move(tempPartial, outputPath, overwrite: true);

            var postHash = ComputeSha256(inputPath);
            if (!string.Equals(preHash, postHash, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"CRITICAL: Source file '{inputPath}' was modified during metadata update operation!");
            }

            progress?.Report(new ProcessingProgress(100, "DOCX metadata updated successfully."));

            return new DocumentEditingResult(
                OutputPath: outputPath,
                IsSuccess: true,
                Limitations: [],
                Message: "Document properties successfully updated.");
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
