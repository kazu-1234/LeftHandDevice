using System.Windows;

namespace LeftHandDeviceApp;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        // 保存されたテーマ設定を起動時に復元
        SettingsWindow.ApplySavedTheme();
    }
}
