// v2.0.13
using System;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Navigation;
using Windows.System;

namespace LeftHandDevice.Views
{
    public sealed partial class InfoPage : Page
    {
        private AppState? _state;
        private UpdateCheckResult? _lastResult;

        public InfoPage()
        {
            InitializeComponent();
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            base.OnNavigatedTo(e);
            _state = e.Parameter as AppState;
            VersionText.Text = Strings.Format("Version_Format", UpdateChecker.CurrentVersion);
        }

        private async void CheckUpdateButton_Click(object sender, RoutedEventArgs e)
        {
            CheckUpdateButton.IsEnabled = false;
            UpdateInfoBar.IsOpen = false;
            InstallUpdateCard.Visibility = Visibility.Collapsed;
            _lastResult = null;

            UpdateCheckResult result = await UpdateChecker.CheckForUpdateAsync();
            _lastResult = result;

            if (_state != null)
            {
                string stamp = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
                _state.Device.SetLastUpdateCheck(stamp);
            }

            UpdateInfoBar.Message = result.Message;
            UpdateInfoBar.IsOpen = true;
            UpdateInfoBar.Severity = result.Status switch
            {
                UpdateCheckStatus.UpdateAvailable => InfoBarSeverity.Informational,
                UpdateCheckStatus.Error => InfoBarSeverity.Error,
                UpdateCheckStatus.NotConfigured => InfoBarSeverity.Informational,
                _ => InfoBarSeverity.Success
            };

            CheckUpdateButton.IsEnabled = true;

            if (result.Status == UpdateCheckStatus.UpdateAvailable)
            {
                if (!string.IsNullOrWhiteSpace(result.DownloadUrl)
                    && !string.IsNullOrWhiteSpace(result.AssetFileName))
                {
                    InstallUpdateCard.Visibility = Visibility.Visible;
                    InstallStatusText.Text = Strings.Format(
                        "Update_DownloadReady",
                        result.LatestVersion ?? string.Empty);
                    InstallUpdateButton.IsEnabled = true;
                }
                else if (!string.IsNullOrWhiteSpace(result.ReleasePageUrl))
                {
                    var dialog = new ContentDialog
                    {
                        Title = Strings.Get("Update_AvailableTitle"),
                        Content = result.Message,
                        PrimaryButtonText = Strings.Get("Update_OpenRelease"),
                        CloseButtonText = Strings.Get("Common_Cancel"),
                        DefaultButton = ContentDialogButton.Primary,
                        XamlRoot = XamlRoot
                    };

                    if (await dialog.ShowAsync() == ContentDialogResult.Primary)
                        await Launcher.LaunchUriAsync(new Uri(result.ReleasePageUrl));
                }
            }
        }

        private async void InstallUpdateButton_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult?.DownloadUrl == null || _lastResult.AssetFileName == null)
                return;

            InstallUpdateButton.IsEnabled = false;
            InstallStatusText.Text = Strings.Get("Update_Preparing");

            try
            {
                var progress = new Progress<string>(msg => InstallStatusText.Text = msg);
                string message = await UpdateInstallerService.DownloadAndInstallAsync(
                    _lastResult.DownloadUrl,
                    _lastResult.AssetFileName,
                    progress);
                InstallStatusText.Text = message;
            }
            catch (Exception ex)
            {
                InstallStatusText.Text = Strings.Format("Update_Failed", ex.Message);
                InstallUpdateButton.IsEnabled = true;
            }
        }
    }
}
