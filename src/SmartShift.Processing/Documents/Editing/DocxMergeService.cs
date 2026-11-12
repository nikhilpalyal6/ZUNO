using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using SmartShift.Core.IO;
using SmartShift.Processing;

namespace SmartShift.Processing.Documents.Editing;

public sealed class DocxMergeService : IDocxMergeService
{
    private readonly IOutputPathAllocator _allocator;

    public DocxMergeService(IOutputPathAllocator allocator)
    {
        _allocator = allocator ?? throw new ArgumentNullException(nameof(allocator));
    }

    public async Task<DocxMergeResult> MergeAsync(
        IReadOnlyList<string> inputPaths,
        string outputPath,
        DocxMergeOptions? options = null,
        IProgress<ProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(inputPaths);
        if (inputPaths.Count < 2)
        {
            throw new ArgumentException("At least two DOCX documents are required to merge.", nameof(inputPaths));
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            throw new ArgumentException("Output path must be specified.", nameof(outputPath));
        }

        options ??= new DocxMergeOptions();
        cancellationToken.ThrowIfCancellationRequested();

        return await Task.Run(() => MergeInternal(inputPaths, outputPath, options, progress, cancellationToken), cancellationToken)
            .ConfigureAwait(false);
    }

    private static DocxMergeResult MergeInternal(
        IReadOnlyList<string> inputPaths,
        string outputPath,
        DocxMergeOptions options,
        IProgress<ProcessingProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        // 1. Pre-validation and pre-hash calculation for source non-destruction verification
        progress?.Report(new ProcessingProgress(5, "Validating input DOCX documents…"));
        var preHashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < inputPaths.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var path = inputPaths[i];
            if (!File.Exists(path))
            {
                throw new FileNotFoundException($"Input DOCX file not found: {path}", path);
            }

            preHashes[path] = ComputeSha256(path);
        }

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

        var limitations = new List<string>();
        int totalParagraphs = 0;
        int totalTables = 0;
        int totalImages = 0;

        try
        {
            progress?.Report(new ProcessingProgress(15, "Initializing merged document container…"));

            // Use first document as base template (retains its styles, core settings, numbering, font table)
            File.Copy(inputPaths[0], tempPartial, overwrite: true);

            using (var targetDoc = WordprocessingDocument.Open(tempPartial, true))
            {
                var mainPart = targetDoc.MainDocumentPart
                    ?? throw new InvalidOperationException("First document does not contain a valid MainDocumentPart.");

                if (mainPart.Document == null)
                {
                    mainPart.Document = new Document();
                }

                var targetBody = mainPart.Document.Body ?? mainPart.Document.AppendChild(new Body());

                totalParagraphs += targetBody.Elements<Paragraph>().Count();
                totalTables += targetBody.Elements<Table>().Count();
                totalImages += mainPart.ImageParts.Count();

                // Detach base document's trailing SectionProperties to convert to paragraph section break if multi-document
                SectionProperties? baseSectPr = targetBody.Elements<SectionProperties>().LastOrDefault();
                if (baseSectPr != null && inputPaths.Count > 1)
                {
                    baseSectPr.Remove();
                    if (options.PreserveSectionProperties)
                    {
                        var sectionBreakP = new Paragraph(new ParagraphProperties(baseSectPr));
                        targetBody.AppendChild(sectionBreakP);
                    }
                    else if (options.AddPageBreakBetweenDocuments)
                    {
                        var pageBreakP = new Paragraph(new Run(new Break { Type = BreakValues.Page }));
                        targetBody.AppendChild(pageBreakP);
                    }
                }
                else if (options.AddPageBreakBetweenDocuments)
                {
                    var pageBreakP = new Paragraph(new Run(new Break { Type = BreakValues.Page }));
                    targetBody.AppendChild(pageBreakP);
                }

                // Append documents 1..N
                for (int i = 1; i < inputPaths.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var currentPath = inputPaths[i];
                    var percent = 20 + (int)((i / (double)inputPaths.Count) * 70);
                    progress?.Report(new ProcessingProgress(percent, $"Merging document {i + 1} of {inputPaths.Count}…"));

                    using var srcDoc = WordprocessingDocument.Open(currentPath, false);
                    var srcMain = srcDoc.MainDocumentPart;
                    if (srcMain?.Document?.Body == null)
                    {
                        limitations.Add($"Document '{Path.GetFileName(currentPath)}' contained no body content.");
                        continue;
                    }

                    // Copy embedded images and track relationship mapping
                    var relIdMap = new Dictionary<string, string>(StringComparer.Ordinal);
                    foreach (var srcImg in srcMain.ImageParts)
                    {
                        var oldRelId = srcMain.GetIdOfPart(srcImg);
                        var targetImg = mainPart.AddImagePart(srcImg.ContentType);
                        using (var srcStream = srcImg.GetStream(FileMode.Open, FileAccess.Read))
                        {
                            targetImg.FeedData(srcStream);
                        }
                        var newRelId = mainPart.GetIdOfPart(targetImg);
                        relIdMap[oldRelId] = newRelId;
                        totalImages++;
                    }

                    var srcElements = srcMain.Document.Body.Elements().ToList();
                    SectionProperties? lastSectPr = null;

                    for (int elemIdx = 0; elemIdx < srcElements.Count; elemIdx++)
                    {
                        var elem = srcElements[elemIdx];

                        if (elem is SectionProperties sp)
                        {
                            lastSectPr = (SectionProperties)sp.CloneNode(true);
                            continue;
                        }

                        var clone = elem.CloneNode(true);

                        // Remap relationship IDs in Drawings and Pictures
                        if (relIdMap.Count > 0)
                        {
                            foreach (var blip in clone.Descendants<A.Blip>())
                            {
                                if (blip.Embed?.Value != null && relIdMap.TryGetValue(blip.Embed.Value, out var newId))
                                {
                                    blip.Embed.Value = newId;
                                }
                            }

                            foreach (var vmlImg in clone.Descendants<V.ImageData>())
                            {
                                if (vmlImg.RelationshipId?.Value != null && relIdMap.TryGetValue(vmlImg.RelationshipId.Value, out var newId))
                                {
                                    vmlImg.RelationshipId.Value = newId;
                                }
                            }
                        }

                        targetBody.AppendChild(clone);

                        if (elem is Paragraph) totalParagraphs++;
                        if (elem is Table) totalTables++;
                    }

                    bool isLastDoc = (i == inputPaths.Count - 1);
                    if (lastSectPr != null)
                    {
                        if (isLastDoc)
                        {
                            targetBody.AppendChild(lastSectPr);
                        }
                        else if (options.PreserveSectionProperties)
                        {
                            var sectionBreakP = new Paragraph(new ParagraphProperties(lastSectPr));
                            targetBody.AppendChild(sectionBreakP);
                        }
                        else if (options.AddPageBreakBetweenDocuments)
                        {
                            var pageBreakP = new Paragraph(new Run(new Break { Type = BreakValues.Page }));
                            targetBody.AppendChild(pageBreakP);
                        }
                    }
                    else if (!isLastDoc && options.AddPageBreakBetweenDocuments)
                    {
                        var pageBreakP = new Paragraph(new Run(new Break { Type = BreakValues.Page }));
                        targetBody.AppendChild(pageBreakP);
                    }
                }

                mainPart.Document.Save();
            }

            progress?.Report(new ProcessingProgress(95, "Finalizing merged document…"));

            // Atomic promotion
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
            File.Move(tempPartial, outputPath, overwrite: true);

            // Verify non-destructive integrity of all source files
            foreach (var kvp in preHashes)
            {
                var postHash = ComputeSha256(kvp.Key);
                if (!string.Equals(kvp.Value, postHash, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"CRITICAL: Source file '{kvp.Key}' was modified during merge operation!");
                }
            }

            progress?.Report(new ProcessingProgress(100, "DOCX merge completed successfully."));

            return new DocxMergeResult(
                OutputPath: outputPath,
                DocumentCount: inputPaths.Count,
                ParagraphCount: totalParagraphs,
                TableCount: totalTables,
                ImageCount: totalImages,
                Limitations: limitations,
                IsSuccess: true);
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
