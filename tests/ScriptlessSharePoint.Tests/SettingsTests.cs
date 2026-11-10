using ScriptlessSharePoint.Settings;

namespace ScriptlessSharePoint.Tests;

public sealed class SettingsTests : IDisposable
{
    private readonly string folder = Path.Combine(Path.GetTempPath(), "ssp-settings-" + Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        if (Directory.Exists(folder))
        {
            Directory.Delete(folder, true);
        }
    }

    [Fact]
    public void Defaults_when_there_is_no_file()
    {
        Assert.False(new SettingsStore(folder).Load().ConfirmChanges);
    }

    [Fact]
    public void Keeps_what_was_saved()
    {
        var store = new SettingsStore(folder);
        store.Save(new AppSettings { ConfirmChanges = true });
        Assert.True(new SettingsStore(folder).Load().ConfirmChanges);
    }

    [Fact]
    public void Defaults_when_the_file_is_broken()
    {
        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, "settings.json"), "{ not json");
        Assert.False(new SettingsStore(folder).Load().ConfirmChanges);
    }
}
