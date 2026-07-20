// v2.0.13
// Inno Setup の setup.exe をダウンロードして起動する更新ヘルパー
using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;

namespace LeftHandDevice
{
    /// <summary>
    /// GitHub Release の Inno setup.exe を Temp に DL して起動する。
    /// folder インストール／単体 exe 差し替えには対応しない。
    /// </summary>
    public static class UpdateInstallerService
    {
        public static async Task<string> DownloadAndInstallAsync(
            string downloadUrl,
            string fileName,
            IProgress<string>? progress = null)
        {
            if (!fileName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                return Strings.Get("Update_SetupRequired");

            // Temp に DL（インストール先の exe を上書きしない）
            string downloadDir = Path.Combine(Path.GetTempPath(), "LeftHandDeviceUpdate");
            Directory.CreateDirectory(downloadDir);
            string targetPath = Path.Combine(downloadDir, fileName);

            progress?.Report(Strings.Get("Update_Downloading"));
            using var client = new HttpClient { Timeout = TimeSpan.FromMinutes(10) };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("LeftHandDevice");

            await using (Stream response = await client.GetStreamAsync(downloadUrl))
            await using (FileStream file = File.Create(targetPath))
            {
                await response.CopyToAsync(file);
            }

            progress?.Report(Strings.Get("Update_LaunchingSetup"));

            Process.Start(new ProcessStartInfo
            {
                FileName = targetPath,
                UseShellExecute = true
            });

            // インストーラが上書きできるようアプリを終了
            try { App.Runtime.ExitApplication(); }
            catch { Application.Current.Exit(); }

            return Strings.Get("Update_SetupStarted");
        }
    }
}
