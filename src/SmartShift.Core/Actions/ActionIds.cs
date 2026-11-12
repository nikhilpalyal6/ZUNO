namespace SmartShift.Core.Actions;

public static class ActionIds
{
    public const string ShowImageInfo = "image.info";
    public const string CompressImage = "image.compress";
    public const string ConvertToWebp = "image.convert.webp";
    public const string ConvertToPng = "image.convert.png";
    public const string ConvertToJpeg = "image.convert.jpeg";
    public const string ResizeImage = "image.resize";
    public const string RemoveImageMetadata = "image.metadata.strip";
    public const string EditImage = "image.edit";

    public const string ShowPdfInfo = "pdf.info";
    public const string CompressPdf = "pdf.compress";
    public const string ExportPdfPages = "pdf.export.pages";
    public const string ExtractPdfText = "pdf.extract.text";
    public const string ConvertPdfToDocx = "pdf.convert.docx";
    public const string ConvertPdfToPptx = "pdf.convert.pptx";
    public const string PreviewPdf = "pdf.preview";
    public const string MergePdf = "pdf.merge";
    public const string SplitPdf = "pdf.split";
    public const string OrganizePdf = "pdf.organize";
    public const string WatermarkPdf = "pdf.watermark";
    public const string EditPdfMetadata = "pdf.metadata";

    public const string ShowDocumentInfo = "doc.info";
    public const string ImportToPresentationStudio = "doc.presentation.import";
    public const string MergeDocx = "docx.merge";
    public const string AddDocxPageNumbers = "docx.pagenumbers";
    public const string AddDocxWatermark = "docx.watermark";
    public const string EditDocxMetadata = "docx.metadata";
    public const string ConvertDocxToPdf = "docx.convert.pdf";
    public const string ConvertDocxToPptx = "docx.convert.pptx";
    public const string ConvertDocxToTxt = "docx.convert.txt";
    public const string ConvertDocxToHtml = "docx.convert.html";
    public const string CompressDocx = "docx.compress";
    public const string ConvertPptxToPdf = "pptx.convert.pdf";
    public const string ExportPptxImages = "pptx.export.images";
    public const string ExtractPptxText = "pptx.extract.text";
    public const string UnifiedConvert = "file.convert";
}

public static class GoalIds
{
    public const string Inspect = "goal.inspect";
    public const string ReduceImageForUpload = "goal.image.reduce-upload";
    public const string ConvertImageForWeb = "goal.image.web";
    public const string ResizeImage = "goal.image.resize";
    public const string EditImage = "goal.image.edit";
    public const string ReducePdfForSharing = "goal.pdf.reduce-share";
    public const string ExportPdfPages = "goal.pdf.export-pages";
    public const string ExtractPdfText = "goal.pdf.extract-text";
    public const string ConvertPdfToDocx = "goal.pdf.convert-docx";
    public const string ConvertPdfToPptx = "goal.pdf.convert-pptx";
    public const string MergePdf = "goal.pdf.merge";
    public const string SplitPdf = "goal.pdf.split";
    public const string OrganizePdf = "goal.pdf.organize";
    public const string WatermarkPdf = "goal.pdf.watermark";
    public const string EditPdfMetadata = "goal.pdf.metadata";
    public const string MergeDocx = "goal.docx.merge";
    public const string AddDocxPageNumbers = "goal.docx.pagenumbers";
    public const string AddDocxWatermark = "goal.docx.watermark";
    public const string EditDocxMetadata = "goal.docx.metadata";
    public const string ConvertDocxToPdf = "goal.docx.convert-pdf";
    public const string ConvertDocxToPptx = "goal.docx.convert-pptx";
    public const string ConvertDocxToTxt = "goal.docx.convert-txt";
    public const string ConvertDocxToHtml = "goal.docx.convert-html";
    public const string CompressDocx = "goal.docx.compress";
    public const string ConvertPptxToPdf = "goal.pptx.convert-pdf";
    public const string ExportPptxImages = "goal.pptx.export-images";
    public const string ExtractPptxText = "goal.pptx.extract-text";
    public const string ImportToPresentation = "goal.doc.presentation.import";
    public const string UnifiedConvert = "goal.file.convert";
}
