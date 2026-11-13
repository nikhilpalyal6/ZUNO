using SmartShift.Core.Settings;

namespace SmartShift.Infrastructure.Settings;

public interface ISettingsService
{
    AppSettings Current { get; }
    string StoragePath { get; }
    event EventHandler<AppSettings>? SettingsChanged;

    AppSettings Load();
    void Save(AppSettings settings);
    AppSettings ResetToDefaults();
}
