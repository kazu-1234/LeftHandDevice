// AppSelectorDialog.xaml.cs
// レジストリ Uninstall キーからインストール済みアプリ一覧を取得し、EXE パスを返す
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.Win32;

namespace LeftHandDevice.Views
{
    /// <summary>インストール済みアプリ選択用の 1 行データ。</summary>
    public sealed class AppItem
    {
        public string Name { get; set; } = "";
        public string ExecutablePath { get; set; } = "";
    }

    public sealed partial class AppSelectorDialog : ContentDialog
    {
        private List<AppItem> _allApps = new();

        /// <summary>選択された実行ファイルのフルパス。未選択時は null。</summary>
        public string? SelectedExecutablePath { get; private set; }

        public AppSelectorDialog()
        {
            InitializeComponent();
            LoadInstalledApps();
        }

        /// <summary>
        /// ダイアログを表示し、選択された EXE パスを返す。
        /// キャンセル時は null。
        /// </summary>
        public static async Task<string?> ShowAsync(XamlRoot xamlRoot)
        {
            var dialog = new AppSelectorDialog
            {
                XamlRoot = xamlRoot
            };
            ContentDialogResult result = await dialog.ShowAsync();
            if (result == ContentDialogResult.Primary)
                return dialog.SelectedExecutablePath;
            return null;
        }

        private void LoadInstalledApps()
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var results = new List<AppItem>();

            void SearchRegistry(RegistryKey baseKey, string keyPath)
            {
                try
                {
                    using var key = baseKey.OpenSubKey(keyPath);
                    if (key == null) return;

                    foreach (string subkeyName in key.GetSubKeyNames())
                    {
                        try
                        {
                            using var subkey = key.OpenSubKey(subkeyName);
                            if (subkey == null) continue;

                            string? displayName = subkey.GetValue("DisplayName") as string;
                            string? installLocation = subkey.GetValue("InstallLocation") as string;
                            string? displayIcon = subkey.GetValue("DisplayIcon") as string;

                            if (string.IsNullOrEmpty(displayName))
                                continue;

                            string? exePath = null;

                            // DisplayIcon が .exe を指していればそれを優先
                            if (!string.IsNullOrEmpty(displayIcon))
                            {
                                int commaIndex = displayIcon.IndexOf(',');
                                string potentialPath = commaIndex > 0
                                    ? displayIcon.Substring(0, commaIndex)
                                    : displayIcon;
                                potentialPath = potentialPath.Trim('"', ' ');
                                if (potentialPath.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
                                    && File.Exists(potentialPath))
                                {
                                    exePath = potentialPath;
                                }
                            }

                            // InstallLocation 直下の exe から推定
                            if (string.IsNullOrEmpty(exePath)
                                && !string.IsNullOrEmpty(installLocation)
                                && Directory.Exists(installLocation))
                            {
                                try
                                {
                                    var exes = Directory.GetFiles(
                                        installLocation, "*.exe", SearchOption.TopDirectoryOnly);
                                    exePath = exes.FirstOrDefault(e =>
                                        e.IndexOf(displayName, StringComparison.OrdinalIgnoreCase) >= 0)
                                        ?? exes.FirstOrDefault();
                                }
                                catch { }
                            }

                            if (!string.IsNullOrEmpty(exePath) && seen.Add(exePath))
                            {
                                results.Add(new AppItem
                                {
                                    Name = displayName,
                                    ExecutablePath = exePath
                                });
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }

            const string uninstallPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            const string wow64UninstallPath =
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall";

            SearchRegistry(Registry.LocalMachine, uninstallPath);
            SearchRegistry(Registry.LocalMachine, wow64UninstallPath);
            SearchRegistry(Registry.CurrentUser, uninstallPath);

            _allApps = results.OrderBy(a => a.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
            AppListView.ItemsSource = _allApps;
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string query = SearchBox.Text?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(query))
            {
                AppListView.ItemsSource = _allApps;
                return;
            }

            AppListView.ItemsSource = _allApps
                .Where(a =>
                    a.Name.Contains(query, StringComparison.OrdinalIgnoreCase)
                    || a.ExecutablePath.Contains(query, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        private void AppListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            IsPrimaryButtonEnabled = AppListView.SelectedItem is AppItem;
        }

        private void ContentDialog_PrimaryButtonClick(
            ContentDialog sender, ContentDialogButtonClickEventArgs args)
        {
            if (AppListView.SelectedItem is AppItem selected)
            {
                SelectedExecutablePath = selected.ExecutablePath;
            }
            else
            {
                // 未選択なら閉じさせない
                args.Cancel = true;
            }
        }
    }
}
