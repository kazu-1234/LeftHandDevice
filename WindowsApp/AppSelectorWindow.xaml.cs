using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace LeftHandDeviceApp
{
    public class AppItem
    {
        public string Name { get; set; }
        public string ExecutablePath { get; set; }
    }

    public partial class AppSelectorWindow : Window
    {
        private List<AppItem> _allApps = new List<AppItem>();

        public string SelectedExecutablePath { get; private set; }

        public AppSelectorWindow()
        {
            InitializeComponent();
            LoadInstalledApps();
        }

        private void LoadInstalledApps()
        {
            var apps = new HashSet<string>();
            var results = new List<AppItem>();

            void SearchRegistry(RegistryKey baseKey, string keyPath)
            {
                try
                {
                    using (var key = baseKey.OpenSubKey(keyPath))
                    {
                        if (key == null) return;
                        foreach (string subkeyName in key.GetSubKeyNames())
                        {
                            using (var subkey = key.OpenSubKey(subkeyName))
                            {
                                if (subkey == null) continue;
                                string displayName = subkey.GetValue("DisplayName") as string;
                                string installLocation = subkey.GetValue("InstallLocation") as string;
                                string displayIcon = subkey.GetValue("DisplayIcon") as string;

                                if (string.IsNullOrEmpty(displayName)) continue;
                                
                                string exePath = null;
                                
                                // Try to extract from DisplayIcon if it's an exe
                                if (!string.IsNullOrEmpty(displayIcon))
                                {
                                    int commaIndex = displayIcon.IndexOf(',');
                                    string potentialPath = commaIndex > 0 ? displayIcon.Substring(0, commaIndex) : displayIcon;
                                    potentialPath = potentialPath.Trim('"', ' ');
                                    if (potentialPath.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) && File.Exists(potentialPath))
                                    {
                                        exePath = potentialPath;
                                    }
                                }
                                
                                // Process InstallLocation
                                if (string.IsNullOrEmpty(exePath) && !string.IsNullOrEmpty(installLocation) && Directory.Exists(installLocation))
                                {
                                    try
                                    {
                                        var exes = Directory.GetFiles(installLocation, "*.exe", SearchOption.TopDirectoryOnly);
                                        exePath = exes.FirstOrDefault(e => e.IndexOf(displayName, StringComparison.OrdinalIgnoreCase) >= 0) ?? exes.FirstOrDefault();
                                    }
                                    catch { }
                                }

                                if (!string.IsNullOrEmpty(exePath) && apps.Add(exePath.ToLower()))
                                {
                                    results.Add(new AppItem { Name = displayName, ExecutablePath = exePath });
                                }
                            }
                        }
                    }
                }
                catch { }
            }

            string uninstallPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            string wow64UninstallPath = @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall";

            SearchRegistry(Registry.LocalMachine, uninstallPath);
            SearchRegistry(Registry.LocalMachine, wow64UninstallPath);
            SearchRegistry(Registry.CurrentUser, uninstallPath);

            _allApps = results.OrderBy(a => a.Name).ToList();
            AppListBox.ItemsSource = _allApps;
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string query = SearchBox.Text.ToLower();
            if (string.IsNullOrWhiteSpace(query))
            {
                AppListBox.ItemsSource = _allApps;
            }
            else
            {
                AppListBox.ItemsSource = _allApps.Where(a => 
                    a.Name.ToLower().Contains(query) || 
                    a.ExecutablePath.ToLower().Contains(query)).ToList();
            }
        }

        private void AppListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectButton.IsEnabled = AppListBox.SelectedItem != null;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            if (AppListBox.SelectedItem is AppItem selected)
            {
                SelectedExecutablePath = selected.ExecutablePath;
                DialogResult = true;
                Close();
            }
        }
    }
}
