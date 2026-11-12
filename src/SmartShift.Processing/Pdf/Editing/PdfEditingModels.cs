namespace SmartShift.Processing.Pdf.Editing;

public sealed record PdfWatermarkOptions(
    string Text,
    double FontSize = 48.0,
    double Opacity = 0.25,
    bool IsDiagonal = true,
    string HexColor = "808080",
    string FontName = "Arial",
    string PageRange = "all"
);

public sealed record PdfMetadata(
    string? Title = null,
    string? Author = null,
    string? Subject = null,
    string? Keywords = null,
    string? Creator = null,
    string? Producer = null
);
