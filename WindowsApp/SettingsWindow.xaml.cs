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
using System.Windows.Media;
using System.Windows.Documents;
using System.Windows.Input;
using System.Threading.Tasks;
using MaterialDesignThemes.Wpf;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace LeftHandDeviceApp
{
    public partial class SettingsWindow : Window
    {
        // アプリのバージョン
        public const string AppVersion = "1.16.0";

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

            // ウィンドウ表示後に自動でアップデート確認を実行
            Loaded += async (s, e) => await PerformUpdateCheck();
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
        private int _selectedIndex = -1;

        private void UpdateReorderList()
        {
            ReorderListGrid.Children.Clear();
            ReorderListGrid.RowDefinitions.Clear();

            for (int i = 0; i < _patterns.Count; i++)
            {
                ReorderListGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(40) }); // 項目の高さを40pxに完全固定
            }

            for (int i = 0; i < _patterns.Count; i++)
            {
                var p = _patterns[i];
                string title = !string.IsNullOrWhiteSpace(p.Name) ? p.Name : $"パターン{i + 1}";

                var itemGrid = new Grid { Background = (_selectedIndex == i) ? new SolidColorBrush(Color.FromArgb(40, 100, 149, 237)) : Brushes.Transparent };
                itemGrid.Tag = p;
                itemGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // テキスト用
                itemGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) }); // ドラッグハンドル用
                
                var dragHandle = new TextBlock
                {
                    Text = "☰",
                    FontSize = 18,
                    FontWeight = FontWeights.Bold,
                    Foreground = (Brush)FindResource("MaterialDesignBodyLight"),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Background = Brushes.Transparent // クリック判定用
                };
                Grid.SetColumn(dragHandle, 1);
                itemGrid.Children.Add(dragHandle);

                var border = new Border { Padding = new Thickness(10, 10, 10, 10), BorderBrush = (Brush)FindResource("MaterialDesignDivider"), BorderThickness = new Thickness(0,0,0,1) };
                var text = new TextBlock { Text = title, VerticalAlignment = VerticalAlignment.Center };
                border.Child = text;
                Grid.SetColumn(border, 0);
                itemGrid.Children.Add(border);
                
                Grid.SetRow(itemGrid, i);
                ReorderListGrid.Children.Add(itemGrid);

                var translate = new TranslateTransform();
                itemGrid.RenderTransform = translate;
                bool isDragging = false;
                Point startMousePos = new Point();
                double itemHeight = 40.0; // ドラッグ計算用変数も40.0に統一

                dragHandle.MouseLeftButtonDown += (s, e) =>
                {
                    _selectedIndex = _patterns.IndexOf(p);
                    foreach (Grid child in ReorderListGrid.Children)
                        child.Background = (child == itemGrid) ? new SolidColorBrush(Color.FromArgb(40, 100, 149, 237)) : Brushes.Transparent;

                    isDragging = true;
                    translate.BeginAnimation(TranslateTransform.YProperty, null);
                    startMousePos = e.GetPosition(ReorderListGrid);
                    dragHandle.CaptureMouse();
                    Panel.SetZIndex(itemGrid, 100);
                    itemGrid.Opacity = 0.8;
                };

                dragHandle.MouseMove += (s, e) =>
                {
                    if (!isDragging) return;
                    Point currentPos = e.GetPosition(ReorderListGrid);
                    double offsetY = currentPos.Y - startMousePos.Y;
                    translate.Y = offsetY;

                    int currentIndex = _patterns.IndexOf(p);
                    int stepsMoved = (int)Math.Round(offsetY / itemHeight);
                    int targetIndex = currentIndex + stepsMoved;
                    targetIndex = Math.Max(0, Math.Min(_patterns.Count - 1, targetIndex));

                    if (targetIndex != currentIndex)
                    {
                        // リスト上のデータ入れ替え（保存はMouseUpまで遅延）
                        _patterns.RemoveAt(currentIndex);
                        _patterns.Insert(targetIndex, p);
                        _selectedIndex = targetIndex;

                        // 各UI要素のGrid.Rowを更新してアニメーション
                        foreach (Grid g in ReorderListGrid.Children)
                        {
                            if (g.Tag is PatternMacroConfig cfg)
                            {
                                int requiredRow = _patterns.IndexOf(cfg);
                                int oldRow = Grid.GetRow(g);
                                if (oldRow != requiredRow)
                                {
                                    Grid.SetRow(g, requiredRow);
                                    if (g != itemGrid && g.RenderTransform is TranslateTransform targetTranslate)
                                    {
                                        targetTranslate.BeginAnimation(TranslateTransform.YProperty, null);
                                        targetTranslate.Y = (oldRow < requiredRow) ? -itemHeight : itemHeight;
                                        targetTranslate.BeginAnimation(
                                            TranslateTransform.YProperty,
                                            new System.Windows.Media.Animation.DoubleAnimation(0, TimeSpan.FromMilliseconds(150))
                                            {
                                                EasingFunction = new System.Windows.Media.Animation.CircleEase
                                                { EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut }
                                            });
                                    }
                                }
                            }
                        }

                        // マウス基準点をオフセット
                        startMousePos.Y += (targetIndex - currentIndex) * itemHeight;
                        translate.Y = currentPos.Y - startMousePos.Y;
                    }
                };

                dragHandle.MouseLeftButtonUp += (s, e) =>
                {
                    if (!isDragging) return;
                    isDragging = false;
                    dragHandle.ReleaseMouseCapture();
                    itemGrid.Opacity = 1.0;
                    Panel.SetZIndex(itemGrid, 0);

                    // まずスナップバックアニメーションを再生
                    var snapAnim = new System.Windows.Media.Animation.DoubleAnimation(0, TimeSpan.FromMilliseconds(50))
                    {
                        EasingFunction = new System.Windows.Media.Animation.CircleEase
                        { EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut }
                    };

                    // アニメーション完了後に保存と反映を実行
                    snapAnim.Completed += (s2, e2) =>
                    {
                        SavePatternsForSettings();
                        if (Application.Current.MainWindow is MainWindow mw)
                            mw.ReloadPatternsFromSettings();
                    };

                    translate.BeginAnimation(TranslateTransform.YProperty, snapAnim);
                };
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

            // 警告音設定の反映
            bool warnSound = true;
            if (File.Exists(SettingsFilePath))
            {
                try
                {
                    var jsonWs = JObject.Parse(File.ReadAllText(SettingsFilePath));
                    if (jsonWs["WarningSound"] != null)
                        warnSound = jsonWs["WarningSound"].Value<bool>();
                }
                catch { }
            }
            WarningSoundToggle.IsChecked = warnSound;

            // 最終確認日時の表示
            if (File.Exists(SettingsFilePath))
            {
                try
                {
                    var jsonLc = JObject.Parse(File.ReadAllText(SettingsFilePath));
                    if (jsonLc["LastUpdateCheck"] != null)
                    {
                        LastCheckTimeText.Text = $"最終確認: {jsonLc["LastUpdateCheck"]}";
                    }
                }
                catch { }
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
                    // 並び替えリストを有効ボタン数に合わせて更新
                    LoadPatternsForSettings();
                    // メイン画面に反映（有効ボタン数変更でカード表示を更新）
                    if (Application.Current.MainWindow is MainWindow mw)
                        mw.ReloadPatternsFromSettings();
                }
            }
        }

        // =============================================
        // アップデート確認
        // =============================================

        /// <summary>
        /// 手動ボタンクリック時のアップデート確認
        /// </summary>
        private async void CheckUpdate_Click(
            object sender, RoutedEventArgs e)
        {
            await PerformUpdateCheck();
        }

        /// <summary>
        /// GitHubリリースから最新バージョンを確認する（共通ロジック）
        /// </summary>
        private async Task PerformUpdateCheck()
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
                        "を使用中です";
                }
            }
            catch (HttpRequestException)
            {
                UpdateStatusText.Text =
                    "サーバーに接続できませんでした。";
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

                // 最終確認日時を保存・表示
                string nowStr = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
                LastCheckTimeText.Text = $"最終確認: {nowStr}";
                try
                {
                    JObject settings;
                    if (File.Exists(SettingsFilePath))
                        settings = JObject.Parse(File.ReadAllText(SettingsFilePath));
                    else
                        settings = new JObject();
                    settings["LastUpdateCheck"] = nowStr;
                    File.WriteAllText(SettingsFilePath, settings.ToString());
                }
                catch { }
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

        // =============================================
        // 警告音トグル変更ハンドラ
        // =============================================
        private void WarningSoundToggle_Changed(object sender, RoutedEventArgs e)
        {
            bool isOn = WarningSoundToggle.IsChecked == true;

            // 設定を保存
            try
            {
                JObject settings;
                if (File.Exists(SettingsFilePath))
                    settings = JObject.Parse(File.ReadAllText(SettingsFilePath));
                else
                    settings = new JObject();
                settings["WarningSound"] = isOn;
                File.WriteAllText(SettingsFilePath, settings.ToString());
            }
            catch { }

            // メインウィンドウに反映
            if (Application.Current.MainWindow is MainWindow mw)
                mw.UpdateWarningSoundSetting(isOn);
        }


    }


}
