namespace SmartShift.Processing.Documents.Editing;

public enum PageNumberPosition
{
    Footer,
    Header
}

public enum PageNumberAlignment
{
    Left,
    Center,
    Right
}

public enum PageNumberFormat
{
    /// <summary>Example: "1"</summary>
    NumberOnly,
    /// <summary>Example: "Page 1"</summary>
    PageNumber,
    /// <summary>Example: "Page 1 of N"</summary>
    PageNumberTotal
}

public sealed record DocxPageNumberOptions(
    PageNumberPosition Position = PageNumberPosition.Footer,
    PageNumberAlignment Alignment = PageNumberAlignment.Center,
    PageNumberFormat Format = PageNumberFormat.PageNumberTotal,
    string FontName = "Calibri",
    double FontSize = 10.0,
    bool ApplyToAllSections = true
);

public sealed record DocxWatermarkOptions(
    string Text,
    double FontSize = 54.0,
    double Opacity = 0.25,
    bool IsDiagonal = true,
    string HexColor = "808080",
    string FontName = "Calibri"
);

public sealed record DocxMetadata(
    string? Title = null,
    string? Subject = null,
    string? Author = null,
    string? Keywords = null,
    string? Comments = null,
    string? Category = null
);

public sealed record DocxMergeOptions(
    bool AddPageBreakBetweenDocuments = true,
    bool PreserveSectionProperties = true
);

public sealed record DocxMergeResult(
    string OutputPath,
    int DocumentCount,
    int ParagraphCount,
    int TableCount,
    int ImageCount,
    IReadOnlyList<string> Limitations,
    bool IsSuccess
);

public sealed record DocumentEditingResult(
    string OutputPath,
    bool IsSuccess,
    IReadOnlyList<string> Limitations,
    string? Message = null
);
