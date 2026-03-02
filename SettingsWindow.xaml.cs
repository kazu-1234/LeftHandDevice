// SettingsWindow.xaml.cs
// v1.0.0
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Windows;
using System.Windows.Controls;
using MaterialDesignThemes.Wpf;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace LeftHandDeviceApp
{
    public partial class SettingsWindow : Window
    {
        // アプリのバージョン
        public const string AppVersion = "1.14.0";

        // プロファイル設定ファイル (MainWindowと共通)
        private static readonly string PatternsFilePath = Path.Combine(Path.GetDirectoryName(Environment.ProcessPath ?? AppDomain.CurrentDomain.BaseDirectory), "app_patterns.json");
        private List<PatternMacroConfig> _patterns = new List<PatternMacroConfig>();

        // GitHubリポジトリ情報
        private const string GitHubOwner = "kazu-1234";
        private const string GitHubRepo = "LeftHandDevice";
        private const string GitHubApiUrl =
            "https://api.github.com/repos/{0}/{1}/releases/latest";

        // テーマ設定保存ファイル
        private static readonly string SettingsFilePath =
            Path.Combine(
                Path.GetDirectoryName(Environment.ProcessPath ?? AppDomain.CurrentDomain.BaseDirectory),
                "app_settings.json");

        public SettingsWindow()
        {
            InitializeComponent();
            VersionText.Text = $"v{AppVersion}";
            LoadSettings();
            LoadPatternsForSettings();
        }

        private void LoadPatternsForSettings()
        {
            if (File.Exists(PatternsFilePath))
            {
                try
                {
                    string json = File.ReadAllText(PatternsFilePath);
                    var loaded = JsonConvert.DeserializeObject<List<PatternMacroConfig>>(json);
                    if (loaded != null) _patterns = loaded;
                }
                catch { }
            }

            // 万一空の場合や不足している場合は、5個デフォルト生成 (MainWindowと同じ仕様)
            if (!File.Exists(PatternsFilePath) && _patterns.Count < 5)
            {
                var existingBtnIds = _patterns.Select(p => p.TriggerParam1).ToList();
                for (int i = 1; i <= 5; i++)
                {
                    if (!existingBtnIds.Contains(i))
                    {
                        var p = new PatternMacroConfig { TriggerType = 0, TriggerParam1 = i, Name = $"ボタン{i}" };
                        p.Steps.Add(new MacroStepConfig { Type = "KEY", Data = ((char)('a' + (i-1))).ToString() });
                        _patterns.Add(p);
                    }
                }
                SavePatternsForSettings();
            }

            if (_patterns.Count == 5 && _patterns.All(p => p.TriggerType == 0))
            {
                _patterns = _patterns.OrderBy(p => p.TriggerParam1).ToList();
            }

            UpdateReorderList();
        }

        private void SavePatternsForSettings()
        {
            try
            {
                string json = JsonConvert.SerializeObject(_patterns, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(PatternsFilePath, json);
            }
            catch { }
        }

        private void UpdateReorderList()
        {
            ReorderListBox.Items.Clear();
            for (int i = 0; i < _patterns.Count; i++)
            {
                var p = _patterns[i];
                string title = !string.IsNullOrWhiteSpace(p.Name) ? p.Name : $"パターン{i + 1}";
                ReorderListBox.Items.Add(title);
            }
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            int idx = ReorderListBox.SelectedIndex;
            if (idx > 0)
            {
                var t = _patterns[idx];
                _patterns[idx] = _patterns[idx - 1];
                _patterns[idx - 1] = t;
                SavePatternsForSettings();
                UpdateReorderList();
                ReorderListBox.SelectedIndex = idx - 1;

                // メインウィンドウが開かれていればリロードさせる
                if (Application.Current.MainWindow is MainWindow mw) mw.ReloadPatternsFromSettings();
            }
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            int idx = ReorderListBox.SelectedIndex;
            if (idx >= 0 && idx < _patterns.Count - 1)
            {
                var t = _patterns[idx];
                _patterns[idx] = _patterns[idx + 1];
                _patterns[idx + 1] = t;
                SavePatternsForSettings();
                UpdateReorderList();
                ReorderListBox.SelectedIndex = idx + 1;

                // メインウィンドウが開かれていればリロードさせる
                if (Application.Current.MainWindow is MainWindow mw) mw.ReloadPatternsFromSettings();
            }
        }

        // =============================================
        // テーマ切り替え
        // =============================================

        /// <summary>
        /// 保存済み設定を読み込んで適用する
        /// </summary>
        private void LoadSettings()
        {
            string theme = "system"; // デフォルト
            int actBtnCount = 5;

            if (File.Exists(SettingsFilePath))
            {
                try
                {
                    var json = JObject.Parse(File.ReadAllText(SettingsFilePath));
                    theme = json["theme"]?.ToString() ?? "system";
                    if (json["ActiveButtonCount"] != null) actBtnCount = json["ActiveButtonCount"].Value<int>();
                }
                catch { }
            }

            // ラジオボタンの状態を反映
            switch (theme)
            {
                case "light": ThemeLight.IsChecked = true; break;
                case "dark": ThemeDark.IsChecked = true; break;
                default: ThemeSystem.IsChecked = true; break;
            }

            // コンボボックスの状態を反映
            foreach (ComboBoxItem item in ActiveButtonCombo.Items)
            {
                if (item.Tag.ToString() == actBtnCount.ToString())
                {
                    ActiveButtonCombo.SelectedItem = item;
                    break;
                }
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
            SaveSettings(themeKey);
        }

        /// <summary>
        /// 設定をJSONに保存する
        /// </summary>
        private void SaveSettings(string themeKey, int? actBtnCount = null)
        {
            try
            {
                JObject settings;
                if (File.Exists(SettingsFilePath)) settings = JObject.Parse(File.ReadAllText(SettingsFilePath));
                else settings = new JObject();

                if (themeKey != null) settings["theme"] = themeKey;
                if (actBtnCount.HasValue) settings["ActiveButtonCount"] = actBtnCount.Value;

                File.WriteAllText(SettingsFilePath, settings.ToString());
            }
            catch { }
        }

        private void ActiveButtonCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (ActiveButtonCombo.SelectedItem is System.Windows.Controls.ComboBoxItem item)
            {
                if (int.TryParse(item.Tag.ToString(), out int val))
                {
                    SaveSettings(null, val);
                }
            }
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
                        $"新しいバージョン v{latestVersion} が公開されています。\n自動でダウンロードしてアップデートしますか？",
                        "アップデート",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Information);

                    if (result == MessageBoxResult.Yes)
                    {
                        if (!string.IsNullOrEmpty(downloadUrl) && downloadUrl.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        {
                            UpdateStatusText.Text = "ダウンロード中...";
                            string exePath = Process.GetCurrentProcess().MainModule.FileName;
                            string exeDir = Path.GetDirectoryName(exePath);
                            string exeName = Path.GetFileName(exePath);
                            string newExePath = Path.Combine(exeDir, "LeftHandDeviceApp_new.exe");
                            string batPath = Path.Combine(exeDir, "update.bat");

                            var exeData = await client.GetByteArrayAsync(downloadUrl);
                            File.WriteAllBytes(newExePath, exeData);

                            string batContent = $@"@echo off
timeout /t 2 /nobreak > nul
del ""{exeName}""
rename ""LeftHandDeviceApp_new.exe"" ""{exeName}""
start """" ""{exeName}""
del ""%~f0""
";
                            File.WriteAllText(batPath, batContent);

                            Process.Start(new ProcessStartInfo(batPath) { UseShellExecute = true, CreateNoWindow = true });
                            Application.Current.Shutdown();
                        }
                        else
                        {
                            // fallback to browser
                            Process.Start(new ProcessStartInfo(releaseUrl) { UseShellExecute = true });
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
