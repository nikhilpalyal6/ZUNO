using SmartShift.Core.Actions;
using SmartShift.Core.Classification;

namespace SmartShift.Processing.Quality;

public sealed record ConversionCapabilityDescriptor(
    string OperationId,
    DetectedFormat SourceFormat,
    DetectedFormat TargetFormat,
    bool SupportsStructuralValidation,
    bool SupportsContentValidation,
    bool SupportsVisualValidation,
    IReadOnlyList<string> KnownLimitations);

public static class ConversionCapabilityMatrix
{
    private static readonly Dictionary<string, ConversionCapabilityDescriptor> Descriptors = new(StringComparer.OrdinalIgnoreCase)
    {
        [ActionIds.ConvertPdfToDocx] = new(
            ActionIds.ConvertPdfToDocx,
            DetectedFormat.Pdf,
            DetectedFormat.Docx,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: false,
            KnownLimitations:
            [
                "WordprocessingML flow layout differs from PDF fixed-position canvas.",
                "Scanned PDFs without OCR are preserved as full-page raster images.",
                "Complex vector illustrations may be rasterized."
            ]),

        [ActionIds.ConvertPdfToPptx] = new(
            ActionIds.ConvertPdfToPptx,
            DetectedFormat.Pdf,
            DetectedFormat.Pptx,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: true,
            KnownLimitations:
            [
                "PDF page layout is reconstructed as slide background or canvas shapes.",
                "Animations, transitions, and slide master styles are not present in PDF sources."
            ]),

        [ActionIds.ExportPdfPages] = new(
            ActionIds.ExportPdfPages,
            DetectedFormat.Pdf,
            DetectedFormat.Png,
            SupportsStructuralValidation: true,
            SupportsContentValidation: false,
            SupportsVisualValidation: true,
            KnownLimitations:
            [
                "Raster output loses selectable text streams.",
                "Resolution is bounded by selected DPI setting."
            ]),

        [ActionIds.ExtractPdfText] = new(
            ActionIds.ExtractPdfText,
            DetectedFormat.Pdf,
            DetectedFormat.Txt,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: false,
            KnownLimitations:
            [
                "Plain text output strips font styles, tables, and graphic elements.",
                "Scanned pages with no text streams return empty text without OCR."
            ]),

        [ActionIds.ConvertDocxToPdf] = new(
            ActionIds.ConvertDocxToPdf,
            DetectedFormat.Docx,
            DetectedFormat.Pdf,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: true,
            KnownLimitations:
            [
                "Local layout engine emulates Word pagination using standard margins and fonts.",
                "Complex SmartArt or nested shapes may be simplified."
            ]),

        [ActionIds.ConvertPptxToPdf] = new(
            ActionIds.ConvertPptxToPdf,
            DetectedFormat.Pptx,
            DetectedFormat.Pdf,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: true,
            KnownLimitations:
            [
                "Embedded media, audio, and video elements cannot be rendered in static PDF.",
                "Transitions and trigger animations are flattened to final slide state."
            ]),

        [ActionIds.ExportPptxImages] = new(
            ActionIds.ExportPptxImages,
            DetectedFormat.Pptx,
            DetectedFormat.Png,
            SupportsStructuralValidation: true,
            SupportsContentValidation: false,
            SupportsVisualValidation: true,
            KnownLimitations:
            [
                "Raster slide image export does not retain selectable text.",
                "Transparency depends on slide background configuration."
            ]),

        [ActionIds.ConvertDocxToTxt] = new(
            ActionIds.ConvertDocxToTxt,
            DetectedFormat.Docx,
            DetectedFormat.Txt,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: false,
            KnownLimitations:
            [
                "All rich formatting, tables, images, and headers are reduced to plain text."
            ]),

        [ActionIds.ExtractPptxText] = new(
            ActionIds.ExtractPptxText,
            DetectedFormat.Pptx,
            DetectedFormat.Txt,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: false,
            KnownLimitations:
            [
                "Slide layout and graphics are omitted."
            ]),

        [ActionIds.ConvertDocxToHtml] = new(
            ActionIds.ConvertDocxToHtml,
            DetectedFormat.Docx,
            DetectedFormat.Html,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: false,
            KnownLimitations:
            [
                "HTML flow layout reflows dynamically across viewport widths.",
                "Embedded images are converted to base64 data URIs."
            ]),

        [ActionIds.MergeDocx] = new(
            ActionIds.MergeDocx,
            DetectedFormat.Docx,
            DetectedFormat.Docx,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: false,
            KnownLimitations:
            [
                "Styles with identical names may be unified to the primary document's definitions."
            ]),

        [ActionIds.AddDocxPageNumbers] = new(
            ActionIds.AddDocxPageNumbers,
            DetectedFormat.Docx,
            DetectedFormat.Docx,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: false,
            KnownLimitations:
            [
                "Field evaluation requires a WordprocessingML compliant viewing application."
            ]),

        [ActionIds.AddDocxWatermark] = new(
            ActionIds.AddDocxWatermark,
            DetectedFormat.Docx,
            DetectedFormat.Docx,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: false,
            KnownLimitations:
            [
                "Watermark is anchored in the header part as a background vector shape."
            ]),

        [ActionIds.EditDocxMetadata] = new(
            ActionIds.EditDocxMetadata,
            DetectedFormat.Docx,
            DetectedFormat.Docx,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: false,
            KnownLimitations:
            [
                "Applies to Open XML package core properties only."
            ]),

        [ActionIds.WatermarkPdf] = new(
            ActionIds.WatermarkPdf,
            DetectedFormat.Pdf,
            DetectedFormat.Pdf,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: true,
            KnownLimitations:
            [
                "Watermark is drawn over content streams using vector text."
            ]),

        [ActionIds.EditPdfMetadata] = new(
            ActionIds.EditPdfMetadata,
            DetectedFormat.Pdf,
            DetectedFormat.Pdf,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: true,
            KnownLimitations:
            [
                "Updates PDF Document Information dictionary; XMP metadata is untouched."
            ]),

        [ActionIds.MergePdf] = new(
            ActionIds.MergePdf,
            DetectedFormat.Pdf,
            DetectedFormat.Pdf,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: true,
            KnownLimitations:
            [
                "Encrypted PDF documents cannot be merged without password removal."
            ]),

        [ActionIds.SplitPdf] = new(
            ActionIds.SplitPdf,
            DetectedFormat.Pdf,
            DetectedFormat.Pdf,
            SupportsStructuralValidation: true,
            SupportsContentValidation: true,
            SupportsVisualValidation: true,
            KnownLimitations:
            [
                "Extracted pages retain original dimensions, orientation, and content streams."
            ])
    };

    public static ConversionCapabilityDescriptor? GetDescriptor(string operationId) =>
        Descriptors.TryGetValue(operationId, out var desc) ? desc : null;

    public static IReadOnlyList<ConversionCapabilityDescriptor> GetAllDescriptors() =>
        Descriptors.Values.ToList();
}
