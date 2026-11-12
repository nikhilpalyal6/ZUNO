using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf.IO;
using SkiaSharp;
using SmartShift.Core.Classification;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace SmartShift.Processing.Quality;

public sealed class StructuralFidelityValidator : IStructuralFidelityValidator
{
    public StructuralMetrics Validate(string filePath, DetectedFormat format)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            return new StructuralMetrics(
                PackageValid: false,
                ElementCount: 0,
                PageOrSlideCount: 0,
                ParagraphCount: 0,
                HeadingCount: 0,
                TableCount: 0,
                ImageCount: 0,
                ShapeCount: 0,
                SectionCount: 0,
                StructuralWarnings: ["Target output file does not exist or is inaccessible."]);
        }

        try
        {
            return format switch
            {
                DetectedFormat.Docx => ValidateDocx(filePath),
                DetectedFormat.Pptx => ValidatePptx(filePath),
                DetectedFormat.Pdf => ValidatePdf(filePath),
                DetectedFormat.Png or DetectedFormat.Jpeg or DetectedFormat.Webp => ValidateImage(filePath),
                DetectedFormat.Html => ValidateHtml(filePath),
                DetectedFormat.Txt or DetectedFormat.Markdown => ValidateText(filePath),
                _ => ValidateGeneric(filePath)
            };
        }
        catch (Exception ex)
        {
            return new StructuralMetrics(
                PackageValid: false,
                ElementCount: 0,
                PageOrSlideCount: 0,
                ParagraphCount: 0,
                HeadingCount: 0,
                TableCount: 0,
                ImageCount: 0,
                ShapeCount: 0,
                SectionCount: 0,
                StructuralWarnings: [$"Structural validation failure: {ex.Message}"]);
        }
    }

    private static StructuralMetrics ValidateDocx(string filePath)
    {
        var warnings = new List<string>();
        try
        {
            using var doc = WordprocessingDocument.Open(filePath, false);
            var mainPart = doc.MainDocumentPart;
            if (mainPart?.Document?.Body == null)
            {
                warnings.Add("DOCX MainDocumentPart or Document Body is missing.");
                return new StructuralMetrics(false, 0, 0, 0, 0, 0, 0, 0, 0, warnings);
            }

            var body = mainPart.Document.Body;
            var paragraphs = body.Elements<W.Paragraph>().ToList();
            var tables = body.Elements<W.Table>().ToList();
            var images = mainPart.ImageParts.Count();
            var sections = body.Descendants<W.SectionProperties>().Count();
            if (sections == 0 && body.Elements<W.SectionProperties>().Any()) sections = 1;

            int headings = 0;
            foreach (var p in paragraphs)
            {
                var style = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (!string.IsNullOrEmpty(style) && (style.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) || style.Equals("Title", StringComparison.OrdinalIgnoreCase)))
                {
                    headings++;
                }
            }

            // Check broken image relationships
            var existingRelIds = new HashSet<string>(mainPart.Parts.Select(p => p.RelationshipId), StringComparer.Ordinal);
            foreach (var imgPart in mainPart.ImageParts)
            {
                existingRelIds.Add(mainPart.GetIdOfPart(imgPart));
            }

            foreach (var blip in body.Descendants<A.Blip>())
            {
                var embed = blip.Embed?.Value;
                if (!string.IsNullOrEmpty(embed) && !existingRelIds.Contains(embed))
                {
                    warnings.Add($"Broken image relationship detected: '{embed}'.");
                }
            }

            int totalElements = paragraphs.Count + tables.Count + images;
            return new StructuralMetrics(
                PackageValid: warnings.Count == 0,
                ElementCount: totalElements,
                PageOrSlideCount: Math.Max(1, sections),
                ParagraphCount: paragraphs.Count,
                HeadingCount: headings,
                TableCount: tables.Count,
                ImageCount: images,
                ShapeCount: 0,
                SectionCount: Math.Max(1, sections),
                StructuralWarnings: warnings);
        }
        catch (Exception ex)
        {
            warnings.Add($"OpenXML package validation failed: {ex.Message}");
            return new StructuralMetrics(false, 0, 0, 0, 0, 0, 0, 0, 0, warnings);
        }
    }

    private static StructuralMetrics ValidatePptx(string filePath)
    {
        var warnings = new List<string>();
        try
        {
            using var doc = PresentationDocument.Open(filePath, false);
            var presPart = doc.PresentationPart;
            if (presPart?.Presentation == null)
            {
                warnings.Add("PPTX PresentationPart or Presentation root is missing.");
                return new StructuralMetrics(false, 0, 0, 0, 0, 0, 0, 0, 0, warnings);
            }

            var slideParts = presPart.SlideParts.ToList();
            int slideCount = slideParts.Count;
            int shapeCount = 0;
            int imageCount = 0;
            int tableCount = 0;

            foreach (var sp in slideParts)
            {
                if (sp.Slide == null) continue;

                var shapes = sp.Slide.Descendants<P.Shape>().Count();
                var pictures = sp.Slide.Descendants<P.Picture>().Count();
                var tables = sp.Slide.Descendants<A.Table>().Count();

                shapeCount += shapes;
                imageCount += pictures;
                tableCount += tables;

                if (shapes == 0 && pictures == 0 && tables == 0)
                {
                    warnings.Add($"Slide '{sp.Uri}' contains no visible shapes, pictures, or tables.");
                }
            }

            int totalElements = shapeCount + imageCount + tableCount;
            return new StructuralMetrics(
                PackageValid: warnings.Count == 0,
                ElementCount: totalElements,
                PageOrSlideCount: slideCount,
                ParagraphCount: 0,
                HeadingCount: 0,
                TableCount: tableCount,
                ImageCount: imageCount,
                ShapeCount: shapeCount,
                SectionCount: 1,
                StructuralWarnings: warnings);
        }
        catch (Exception ex)
        {
            warnings.Add($"Presentation package validation failed: {ex.Message}");
            return new StructuralMetrics(false, 0, 0, 0, 0, 0, 0, 0, 0, warnings);
        }
    }

    private static StructuralMetrics ValidatePdf(string filePath)
    {
        var warnings = new List<string>();
        try
        {
            using var doc = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
            int pageCount = doc.PageCount;
            if (pageCount == 0)
            {
                warnings.Add("PDF document contains 0 pages.");
                return new StructuralMetrics(false, 0, 0, 0, 0, 0, 0, 0, 0, warnings);
            }

            if (doc.SecuritySettings.IsEncrypted)
            {
                warnings.Add("PDF document is encrypted or password-protected.");
            }

            return new StructuralMetrics(
                PackageValid: true,
                ElementCount: pageCount,
                PageOrSlideCount: pageCount,
                ParagraphCount: 0,
                HeadingCount: 0,
                TableCount: 0,
                ImageCount: 0,
                ShapeCount: 0,
                SectionCount: 1,
                StructuralWarnings: warnings);
        }
        catch (Exception ex)
        {
            warnings.Add($"PDF reader failed to load document: {ex.Message}");
            return new StructuralMetrics(false, 0, 0, 0, 0, 0, 0, 0, 0, warnings);
        }
    }

    private static StructuralMetrics ValidateImage(string filePath)
    {
        var warnings = new List<string>();
        try
        {
            using var codec = SKCodec.Create(filePath);
            if (codec == null)
            {
                warnings.Add("Failed to create SkiaSharp codec for image file.");
                return new StructuralMetrics(false, 0, 0, 0, 0, 0, 0, 0, 0, warnings);
            }

            var info = codec.Info;
            if (info.Width <= 0 || info.Height <= 0)
            {
                warnings.Add($"Invalid image dimensions: {info.Width}x{info.Height}.");
                return new StructuralMetrics(false, 0, 0, 0, 0, 0, 0, 0, 0, warnings);
            }

            return new StructuralMetrics(
                PackageValid: true,
                ElementCount: 1,
                PageOrSlideCount: 1,
                ParagraphCount: 0,
                HeadingCount: 0,
                TableCount: 0,
                ImageCount: 1,
                ShapeCount: 0,
                SectionCount: 1,
                StructuralWarnings: warnings);
        }
        catch (Exception ex)
        {
            warnings.Add($"Image validation failed: {ex.Message}");
            return new StructuralMetrics(false, 0, 0, 0, 0, 0, 0, 0, 0, warnings);
        }
    }

    private static StructuralMetrics ValidateHtml(string filePath)
    {
        var warnings = new List<string>();
        try
        {
            var text = File.ReadAllText(filePath);
            bool hasHtml = text.Contains("<html", StringComparison.OrdinalIgnoreCase) || text.Contains("<body", StringComparison.OrdinalIgnoreCase);
            if (!hasHtml)
            {
                warnings.Add("HTML document is missing standard HTML root elements.");
            }

            return new StructuralMetrics(
                PackageValid: hasHtml,
                ElementCount: 1,
                PageOrSlideCount: 1,
                ParagraphCount: 0,
                HeadingCount: 0,
                TableCount: text.Contains("<table", StringComparison.OrdinalIgnoreCase) ? 1 : 0,
                ImageCount: text.Contains("<img", StringComparison.OrdinalIgnoreCase) ? 1 : 0,
                ShapeCount: 0,
                SectionCount: 1,
                StructuralWarnings: warnings);
        }
        catch (Exception ex)
        {
            warnings.Add($"HTML validation failed: {ex.Message}");
            return new StructuralMetrics(false, 0, 0, 0, 0, 0, 0, 0, 0, warnings);
        }
    }

    private static StructuralMetrics ValidateText(string filePath)
    {
        var warnings = new List<string>();
        try
        {
            var text = File.ReadAllText(filePath);
            return new StructuralMetrics(
                PackageValid: true,
                ElementCount: 1,
                PageOrSlideCount: 1,
                ParagraphCount: text.Split('\n', StringSplitOptions.RemoveEmptyEntries).Length,
                HeadingCount: 0,
                TableCount: 0,
                ImageCount: 0,
                ShapeCount: 0,
                SectionCount: 1,
                StructuralWarnings: warnings);
        }
        catch (Exception ex)
        {
            warnings.Add($"Text validation failed: {ex.Message}");
            return new StructuralMetrics(false, 0, 0, 0, 0, 0, 0, 0, 0, warnings);
        }
    }

    private static StructuralMetrics ValidateGeneric(string filePath)
    {
        var info = new FileInfo(filePath);
        return new StructuralMetrics(
            PackageValid: info.Exists && info.Length > 0,
            ElementCount: 1,
            PageOrSlideCount: 1,
            ParagraphCount: 0,
            HeadingCount: 0,
            TableCount: 0,
            ImageCount: 0,
            ShapeCount: 0,
            SectionCount: 1,
            StructuralWarnings: info.Length == 0 ? ["Output file is 0 bytes."] : []);
    }
}
