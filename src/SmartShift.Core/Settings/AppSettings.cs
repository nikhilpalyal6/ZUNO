namespace SmartShift.Core.Settings;

public enum QualityValidationLevel
{
    Basic = 0,
    Standard = 1,
    Deep = 2
}

public sealed class AppSettings
{
    public string? DefaultOutputDirectory { get; set; }
    public string OutputNamingPattern { get; set; } = "{name}_{action}_{timestamp}";
    public int DefaultImageQuality { get; set; } = 85;
    public int DefaultPdfDpi { get; set; } = 150;
    public QualityValidationLevel DefaultValidationLevel { get; set; } = QualityValidationLevel.Standard;
    public int MaxConcurrentOperations { get; set; } = 4;
    public bool ShowFirstRunTips { get; set; } = true;
    public bool NotifyOnCompletion { get; set; } = true;
    public string Theme { get; set; } = "Dark";

    public AppSettings Clone()
    {
        return new AppSettings
        {
            DefaultOutputDirectory = DefaultOutputDirectory,
            OutputNamingPattern = OutputNamingPattern,
            DefaultImageQuality = DefaultImageQuality,
            DefaultPdfDpi = DefaultPdfDpi,
            DefaultValidationLevel = DefaultValidationLevel,
            MaxConcurrentOperations = MaxConcurrentOperations,
            ShowFirstRunTips = ShowFirstRunTips,
            NotifyOnCompletion = NotifyOnCompletion,
            Theme = Theme
        };
    }
}
