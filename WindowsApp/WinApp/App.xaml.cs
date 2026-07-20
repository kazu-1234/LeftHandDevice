// v2.0.0
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using System;
using System.Diagnostics;

namespace LeftHandDevice
{
    public partial class App : Application
    {
        private AppRuntime? _runtime;

        internal static AppRuntime Runtime =>
            (Current as App)?._runtime
            ?? throw new InvalidOperationException("App runtime is not initialized.");

        public App()
        {
            InitializeComponent();
            // DeviceService 等のプロセス寿命管理のため明示終了まで落とさない
            DispatcherShutdownMode = DispatcherShutdownMode.OnExplicitShutdown;
            UnhandledException += App_UnhandledException;
        }

        private void App_UnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
        {
            Debug.WriteLine(e.Exception);
        }

        protected override void OnLaunched(LaunchActivatedEventArgs args)
        {
            try
            {
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(AppContext.BaseDirectory, "startup.log"),
                    $"[{DateTime.Now:HH:mm:ss.fff}] OnLaunched begin{Environment.NewLine}");
            }
            catch { }

            UpdateChecker.LatestReleaseApiUrl =
                "https://api.github.com/repos/kazu-1234/LeftHandDevice/releases/latest";

            var settings = Settings.Load();

            bool launchInBackground = HasCommandLineArg("--background");
            bool requestInteractiveShow = !launchInBackground;

            if (!SingleInstanceManager.TryBecomePrimaryInstance(requestInteractiveShow))
            {
                try
                {
                    System.IO.File.AppendAllText(
                        System.IO.Path.Combine(AppContext.BaseDirectory, "startup.log"),
                        $"[{DateTime.Now:HH:mm:ss.fff}] secondary instance -> Exit{Environment.NewLine}");
                }
                catch { }
                Exit();
                return;
            }

            var dq = DispatcherQueue.GetForCurrentThread();
            _runtime = new AppRuntime(this, settings, dq);
            try
            {
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(AppContext.BaseDirectory, "startup.log"),
                    $"[{DateTime.Now:HH:mm:ss.fff}] runtime.Start{Environment.NewLine}");
            }
            catch { }
            _runtime.Start(launchInBackground, requestInteractiveShow);
            try
            {
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(AppContext.BaseDirectory, "startup.log"),
                    $"[{DateTime.Now:HH:mm:ss.fff}] OnLaunched end{Environment.NewLine}");
            }
            catch { }
        }

        private static bool HasCommandLineArg(string arg)
        {
            foreach (string item in Environment.GetCommandLineArgs())
            {
                if (string.Equals(item, arg, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }
    }
}
