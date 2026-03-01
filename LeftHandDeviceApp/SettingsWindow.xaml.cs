// SettingsWindow.xaml.cs
// v1.0.0
using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Windows;
using MaterialDesignThemes.Wpf;
using Newtonsoft.Json.Linq;

namespace LeftHandDeviceApp
{
    public partial class SettingsWindow : Window
    {
        // アプリのバージョン
        public const string AppVersion = "1.13.0";

        // GitHubリポジトリ情報
        private const string GitHubOwner = "kazu-1234";
        private const string GitHubRepo = "LeftHandDevice";
        private const string GitHubApiUrl =
            "https://api.github.com/repos/{0}/{1}/releases/latest";

        // テーマ設定保存ファイル
        private static readonly string SettingsFilePath =
            Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "app_settings.json");

        public SettingsWindow()
        {
            InitializeComponent();
            VersionText.Text = $"v{AppVersion}";
            LoadThemeSetting();
        }

        // =============================================
        // テーマ切り替え
        // =============================================

        /// <summary>
        /// 保存済みテーマ設定を読み込んで適用する
        /// </summary>
        private void LoadThemeSetting()
        {
            string theme = "system"; // デフォルト
            if (File.Exists(SettingsFilePath))
            {
                try
                {
                    var json = JObject.Parse(
                        File.ReadAllText(SettingsFilePath));
                    theme = json["theme"]?.ToString() ?? "system";
                }
                catch { }
            }

            // ラジオボタンの状態を反映（イベント発火でテーマも適用される）
            switch (theme)
            {
                case "light":
                    ThemeLight.IsChecked = true;
                    break;
                case "dark":
                    ThemeDark.IsChecked = true;
                    break;
                default:
                    ThemeSystem.IsChecked = true;
                    break;
            }
        }

        /// <summary>
        /// テーマ変更イベントハンドラ
        /// </summary>
        private void Theme_Changed(
            object sender, RoutedEventArgs e)
        {
            string themeKey;
            BaseTheme baseTheme;

            if (ThemeDark.IsChecked == true)
            {
                themeKey = "dark";
                baseTheme = BaseTheme.Dark;
            }
            else if (ThemeLight.IsChecked == true)
            {
                themeKey = "light";
                baseTheme = BaseTheme.Light;
            }
            else
            {
                themeKey = "system";
                baseTheme = BaseTheme.Inherit;
            }

            // テーマを適用
            var paletteHelper = new PaletteHelper();
            var currentTheme = paletteHelper.GetTheme();
            currentTheme.SetBaseTheme(baseTheme);
            paletteHelper.SetTheme(currentTheme);

            // 設定を保存
            SaveThemeSetting(themeKey);
        }

        /// <summary>
        /// テーマ設定をJSONに保存する
        /// </summary>
        private void SaveThemeSetting(string themeKey)
        {
            try
            {
                JObject settings;
                if (File.Exists(SettingsFilePath))
                {
                    settings = JObject.Parse(
                        File.ReadAllText(SettingsFilePath));
                }
                else
                {
                    settings = new JObject();
                }
                settings["theme"] = themeKey;
                File.WriteAllText(
                    SettingsFilePath,
                    settings.ToString());
            }
            catch { }
        }

        // =============================================
        // アップデート確認
        // =============================================

        /// <summary>
        /// GitHubリリースから最新バージョンを確認する
        /// </summary>
        private async void CheckUpdate_Click(
            object sender, RoutedEventArgs e)
        {
            CheckUpdateButton.IsEnabled = false;
            UpdateProgress.Visibility = Visibility.Visible;
            UpdateStatusText.Text = "確認中...";

            try
            {
                using var client = new HttpClient();
                client.DefaultRequestHeaders.UserAgent.ParseAdd(
                    "LeftHandDeviceApp");

                string url = string.Format(
                    GitHubApiUrl, GitHubOwner, GitHubRepo);
                var response = await client.GetStringAsync(url);
                var release = JObject.Parse(response);

                string latestTag =
                    release["tag_name"]?.ToString() ?? "";
                string latestVersion =
                    latestTag.TrimStart('v', 'V');
                string releaseUrl =
                    release["html_url"]?.ToString() ?? "";

                // ダウンロードURL（最初のアセット）
                string downloadUrl = "";
                var assets = release["assets"] as JArray;
                if (assets != null && assets.Count > 0)
                {
                    downloadUrl =
                        assets[0]?["browser_download_url"]?
                        .ToString() ?? "";
                }

                // バージョン比較
                if (IsNewerVersion(latestVersion, AppVersion))
                {
                    UpdateStatusText.Text =
                        $"新バージョン v{latestVersion}" +
                        " が利用可能です！";
                    var result = MessageBox.Show(
                        $"新しいバージョン v{latestVersion}" +
                        " が公開されています。\n" +
                        "ダウンロードページを開きますか？",
                        "アップデート",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Information);
                    if (result == MessageBoxResult.Yes)
                    {
                        // ダウンロードURLまたはリリースページを開く
                        string openUrl =
                            !string.IsNullOrEmpty(downloadUrl)
                            ? downloadUrl
                            : releaseUrl;
                        if (!string.IsNullOrEmpty(openUrl))
                        {
                            Process.Start(
                                new ProcessStartInfo(openUrl)
                                {
                                    UseShellExecute = true
                                });
                        }
                    }
                }
                else
                {
                    UpdateStatusText.Text =
                        $"最新バージョン (v{AppVersion}) " +
                        "を使用中です ✔";
                }
            }
            catch (HttpRequestException)
            {
                UpdateStatusText.Text =
                    "サーバーに接続できませんでした。" +
                    "リポジトリが公開されているか確認してください。";
            }
            catch (Exception ex)
            {
                UpdateStatusText.Text =
                    $"確認中にエラーが発生しました: " +
                    $"{ex.Message}";
            }
            finally
            {
                CheckUpdateButton.IsEnabled = true;
                UpdateProgress.Visibility = Visibility.Collapsed;
            }
        }

        /// <summary>
        /// バージョン文字列比較（latest > current なら true）
        /// </summary>
        private bool IsNewerVersion(
            string latest, string current)
        {
            try
            {
                var latestVer = new Version(latest);
                var currentVer = new Version(current);
                return latestVer > currentVer;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 起動時にテーマを復元するための静的メソッド
        /// </summary>
        public static void ApplySavedTheme()
        {
            string theme = "system";
            if (File.Exists(SettingsFilePath))
            {
                try
                {
                    var json = JObject.Parse(
                        File.ReadAllText(SettingsFilePath));
                    theme = json["theme"]?.ToString() ?? "system";
                }
                catch { }
            }

            BaseTheme baseTheme = theme switch
            {
                "light" => BaseTheme.Light,
                "dark" => BaseTheme.Dark,
                _ => BaseTheme.Inherit
            };

            var paletteHelper = new PaletteHelper();
            var currentTheme = paletteHelper.GetTheme();
            currentTheme.SetBaseTheme(baseTheme);
            paletteHelper.SetTheme(currentTheme);
        }

        private void Close_Click(
            object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
