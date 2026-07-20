// v2.0.0
using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace LeftHandDevice
{
    public class Settings
    {
        /// <summary>デフォルトはシステム連動。</summary>
        public AppThemePreference ThemePreference { get; set; } = AppThemePreference.System;

        public int WindowWidth { get; set; } = 960;
        public int WindowHeight { get; set; } = 700;

        /// <summary>未保存時は -1。次回起動で位置を復元する。</summary>
        public int WindowX { get; set; } = -1;

        /// <summary>未保存時は -1。次回起動で位置を復元する。</summary>
        public int WindowY { get; set; } = -1;

        /// <summary>前回終了時に最大化されていたか。</summary>
        public bool WindowMaximized { get; set; }

        private static string SettingsFilePath =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "LeftHandDevice",
                "settings.json");

        /// <summary>旧 WPF 版の exe 隣 app_settings.json（テーマ移行用）。</summary>
        private static string LegacyAppSettingsPath
        {
            get
            {
                string baseDir = Path.GetDirectoryName(Environment.ProcessPath)
                    ?? AppDomain.CurrentDomain.BaseDirectory;
                return Path.Combine(baseDir, "app_settings.json");
            }
        }

        public static Settings Load()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    string json = File.ReadAllText(SettingsFilePath);
                    var settings = JsonConvert.DeserializeObject<Settings>(json);
                    if (settings != null)
                        return settings;
                }
            }
            catch
            {
                // 破損時はデフォルト
            }

            var defaults = new Settings();
            defaults.TryMigrateThemeFromLegacy();
            return defaults;
        }

        /// <summary>旧 app_settings.json の theme を初回のみ取り込む。</summary>
        private void TryMigrateThemeFromLegacy()
        {
            try
            {
                if (!File.Exists(LegacyAppSettingsPath))
                    return;

                var json = JObject.Parse(File.ReadAllText(LegacyAppSettingsPath));
                string? theme = json["theme"]?.ToString();
                ThemePreference = theme switch
                {
                    "light" => AppThemePreference.Light,
                    "dark" => AppThemePreference.Dark,
                    _ => AppThemePreference.System
                };
            }
            catch
            {
                // 移行失敗は無視
            }
        }

        public void Save()
        {
            try
            {
                string? dir = Path.GetDirectoryName(SettingsFilePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                string json = JsonConvert.SerializeObject(this, Formatting.Indented);
                File.WriteAllText(SettingsFilePath, json);
            }
            catch
            {
                // 保存失敗は無視
            }
        }
    }
}
