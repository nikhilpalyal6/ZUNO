using System.Text.Json;
using SmartShift.Core.Settings;

namespace SmartShift.Infrastructure.Settings;

public sealed class SettingsService : ISettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    private readonly string _storagePath;
    private readonly object _lock = new();
    private AppSettings _current;

    public event EventHandler<AppSettings>? SettingsChanged;

    public SettingsService(string? storagePath = null)
    {
        _storagePath = storagePath ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ZUNO",
            "settings.json");

        _current = Load();
    }

    public AppSettings Current
    {
        get
        {
            lock (_lock)
            {
                return _current.Clone();
            }
        }
    }

    public string StoragePath => _storagePath;

    public AppSettings Load()
    {
        lock (_lock)
        {
            if (!File.Exists(_storagePath))
            {
                _current = new AppSettings();
                return _current.Clone();
            }

            try
            {
                var json = File.ReadAllText(_storagePath);
                var loaded = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions);
                if (loaded != null)
                {
                    Sanitize(loaded);
                    _current = loaded;
                    return _current.Clone();
                }
            }
            catch
            {
                // Fallback to safe defaults if settings file is corrupted or unreadable
            }

            _current = new AppSettings();
            return _current.Clone();
        }
    }

    public void Save(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        lock (_lock)
        {
            var toSave = settings.Clone();
            Sanitize(toSave);

            var dir = Path.GetDirectoryName(_storagePath);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var tempPath = _storagePath + ".tmp";
            var json = JsonSerializer.Serialize(toSave, JsonOptions);
            File.WriteAllText(tempPath, json);
            File.Move(tempPath, _storagePath, overwrite: true);

            _current = toSave;
        }

        SettingsChanged?.Invoke(this, Current);
    }

    public AppSettings ResetToDefaults()
    {
        var defaults = new AppSettings();
        Save(defaults);
        return defaults;
    }

    private static void Sanitize(AppSettings settings)
    {
        if (settings.DefaultImageQuality is < 10 or > 100)
        {
            settings.DefaultImageQuality = Math.Clamp(settings.DefaultImageQuality, 10, 100);
            if (settings.DefaultImageQuality == 0) settings.DefaultImageQuality = 85;
        }

        if (settings.DefaultPdfDpi is < 72 or > 600)
        {
            settings.DefaultPdfDpi = Math.Clamp(settings.DefaultPdfDpi, 72, 600);
            if (settings.DefaultPdfDpi == 0) settings.DefaultPdfDpi = 150;
        }

        if (settings.MaxConcurrentOperations is < 1 or > 16)
        {
            settings.MaxConcurrentOperations = Math.Clamp(settings.MaxConcurrentOperations, 1, 16);
            if (settings.MaxConcurrentOperations == 0) settings.MaxConcurrentOperations = 4;
        }

        if (string.IsNullOrWhiteSpace(settings.OutputNamingPattern))
        {
            settings.OutputNamingPattern = "{name}_{action}_{timestamp}";
        }

        if (string.IsNullOrWhiteSpace(settings.Theme))
        {
            settings.Theme = "Dark";
        }
    }
}
