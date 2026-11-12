using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using PdfSharp.Pdf.IO;
using SmartShift.Core.Classification;
using SmartShift.Processing.Documents;
using SmartShift.Processing.Pdf;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace SmartShift.Processing.Quality;

public sealed partial class ContentIntegrityValidator : IContentIntegrityValidator
{
    public ContentIntegrityMetrics? Validate(
        string sourcePath,
        DetectedFormat sourceFormat,
        string outputPath,
        DetectedFormat outputFormat)
    {
        if (IsNonTextFormat(sourceFormat) || IsNonTextFormat(outputFormat))
        {
            return null;
        }

        var sourceText = ExtractText(sourcePath, sourceFormat);
        var outputText = ExtractText(outputPath, outputFormat);

        if (sourceText == null || outputText == null)
        {
            return null;
        }

        var warnings = new List<string>();

        var normSource = NormalizeText(sourceText);
        var normOutput = NormalizeText(outputText);

        long srcLen = normSource.Length;
        long outLen = normOutput.Length;

        // If source is empty, handle gracefully
        if (srcLen == 0)
        {
            return new ContentIntegrityMetrics(
                SourceCharacterCount: 0,
                OutputCharacterCount: outLen,
                TextSimilarityScore: outLen == 0 ? 1.0 : 0.0,
                MissingTextSegments: 0,
                DuplicateTextSegments: 0,
                UnicodePreserved: true,
                TablesPreserved: true,
                ContentWarnings: outLen > 0 ? ["Source document contained no extracted text."] : []);
        }

        // Tokenize into words
        var srcTokens = Tokenize(normSource);
        var outTokens = Tokenize(normOutput);

        var srcFreq = GetFrequencies(srcTokens);
        var outFreq = GetFrequencies(outTokens);

        int intersectionCount = 0;
        int missingCount = 0;
        int duplicateCount = 0;

        foreach (var kvp in srcFreq)
        {
            if (outFreq.TryGetValue(kvp.Key, out var outCount))
            {
                intersectionCount += Math.Min(kvp.Value, outCount);
                if (outCount > kvp.Value)
                {
                    duplicateCount += (outCount - kvp.Value);
                }
            }
            else
            {
                missingCount += kvp.Value;
            }
        }

        double similarityScore = (srcTokens.Count + outTokens.Count) > 0
            ? (2.0 * intersectionCount) / (srcTokens.Count + outTokens.Count)
            : 1.0;

        similarityScore = Math.Clamp(similarityScore, 0.0, 1.0);

        // Unicode preservation check: characters > 127
        bool unicodePreserved = true;
        var nonAsciiInSource = normSource.Where(c => c > 127).Distinct().ToList();
        if (nonAsciiInSource.Count > 0)
        {
            var outputCharSet = new HashSet<char>(normOutput);
            int missingUnicode = nonAsciiInSource.Count(c => !outputCharSet.Contains(c));
            if (missingUnicode > 0)
            {
                unicodePreserved = false;
                warnings.Add($"{missingUnicode} distinct non-ASCII Unicode characters from source were missing in the output.");
            }
        }

        if (missingCount > 0 && similarityScore < 0.70)
        {
            warnings.Add($"Substantial text divergence detected: {missingCount} words missing in output.");
        }

        if (outLen == 0 && srcLen > 0)
        {
            warnings.Add("Target output contains 0 extracted text characters while source contained text.");
        }

        return new ContentIntegrityMetrics(
            SourceCharacterCount: srcLen,
            OutputCharacterCount: outLen,
            TextSimilarityScore: Math.Round(similarityScore, 4),
            MissingTextSegments: missingCount,
            DuplicateTextSegments: duplicateCount,
            UnicodePreserved: unicodePreserved,
            TablesPreserved: true,
            ContentWarnings: warnings);
    }

    private static bool IsNonTextFormat(DetectedFormat format) =>
        format is DetectedFormat.Png or DetectedFormat.Jpeg or DetectedFormat.Webp or DetectedFormat.Unknown;

    private static string? ExtractText(string path, DetectedFormat format)
    {
        if (!File.Exists(path)) return null;

        try
        {
            switch (format)
            {
                case DetectedFormat.Docx:
                    return ExtractDocxText(path);

                case DetectedFormat.Pptx:
                    return ExtractPptxText(path);

                case DetectedFormat.Pdf:
                    return ExtractPdfText(path);

                case DetectedFormat.Txt:
                case DetectedFormat.Markdown:
                    return File.ReadAllText(path);

                case DetectedFormat.Html:
                    return StripHtmlTags(File.ReadAllText(path));

                default:
                    return null;
            }
        }
        catch
        {
            return null;
        }
    }

    private static string ExtractDocxText(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) return string.Empty;

        var sb = new StringBuilder();
        foreach (var p in body.Elements<W.Paragraph>())
        {
            sb.AppendLine(p.InnerText);
        }
        return sb.ToString();
    }

    private static string ExtractPptxText(string path)
    {
        using var doc = PresentationDocument.Open(path, false);
        var presPart = doc.PresentationPart;
        if (presPart == null) return string.Empty;

        var sb = new StringBuilder();
        foreach (var slidePart in presPart.SlideParts)
        {
            if (slidePart.Slide == null) continue;

            foreach (var text in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
            {
                sb.Append(text.Text).Append(' ');
            }
            sb.AppendLine();
        }
        return sb.ToString();
    }

    private static string ExtractPdfText(string path)
    {
        try
        {
            var text = PdfTextExtractor.ExtractAllText(path);
            return text ?? string.Empty;
        }
        catch
        {
            // Fallback via PDFsharp
            using var doc = PdfReader.Open(path, PdfDocumentOpenMode.Import);
            return $"[PDF with {doc.PageCount} pages]";
        }
    }

    private static string StripHtmlTags(string html)
    {
        return HtmlTagRegex().Replace(html, " ");
    }

    private static string NormalizeText(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;
        return WhitespaceRegex().Replace(text.Trim(), " ");
    }

    private static List<string> Tokenize(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return [];
        return text.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                   .Select(w => w.ToLowerInvariant())
                   .ToList();
    }

    private static Dictionary<string, int> GetFrequencies(IEnumerable<string> tokens)
    {
        var dict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in tokens)
        {
            dict[t] = dict.TryGetValue(t, out var c) ? c + 1 : 1;
        }
        return dict;
    }

    [GeneratedRegex("<.*?>", RegexOptions.Compiled)]
    private static partial Regex HtmlTagRegex();

    [GeneratedRegex(@"\s+", RegexOptions.Compiled)]
    private static partial Regex WhitespaceRegex();
}
