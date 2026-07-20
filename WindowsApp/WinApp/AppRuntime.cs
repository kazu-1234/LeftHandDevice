// v2.0.10
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinRT.Interop;

namespace LeftHandDevice
{
    /// <summary>
    /// プロセス寿命（二重起動イベント・DeviceService・MainWindow）を管理する。
    /// トレイ常駐・自動起動は実装しない。ウィンドウを閉じたらアプリ終了。
    /// </summary>
    public sealed class AppRuntime : IDisposable
    {
        private readonly Application _app;
        private readonly Settings _settings;
        private readonly AppState _appState;
        private readonly DeviceService _device;
        private readonly DispatcherQueue _dispatcherQueue;

        private MainWindow? _mainWindow;
        private CancellationTokenSource? _listenerCts;
        private int _exitStarted; // Interlocked: 0/1
        private bool _isExitingProcess;

        public AppRuntime(Application app, Settings settings, DispatcherQueue dispatcherQueue)
        {
            _app = app;
            _settings = settings;
            _dispatcherQueue = dispatcherQueue;
            _device = new DeviceService(dispatcherQueue);
            _appState = new AppState(_settings, _device);
        }

        public AppState AppState => _appState;
        public Settings Settings => _settings;
        public DeviceService Device => _device;
        public bool IsExitingProcess => _isExitingProcess;

        public void Start(bool launchInBackground, bool requestInteractiveShow)
        {
            ThemeService.Initialize(_settings.ThemePreference);
            StartListeners();

            _ = launchInBackground;
            _ = requestInteractiveShow;
            ShowOrCreateMainWindowCore();
        }

        public void ShowOrCreateMainWindow(string? pageTag = null)
        {
            if (_isExitingProcess)
                return;

            if (_dispatcherQueue.HasThreadAccess)
                ShowOrCreateMainWindowCore(pageTag);
            else
                _dispatcherQueue.TryEnqueue(() => ShowOrCreateMainWindowCore(pageTag));
        }

        private void ShowOrCreateMainWindowCore(string? pageTag = null)
        {
            if (_isExitingProcess)
                return;

            if (_mainWindow != null)
            {
                BringWindowToForeground(_mainWindow);
                if (pageTag != null)
                    _mainWindow.NavigateToPageTag(pageTag);
                return;
            }

            _mainWindow = new MainWindow(this);
            _mainWindow.Closed += MainWindow_Closed;
            _mainWindow.PrepareAndActivate(pageTag);
        }

        /// <summary>× 押下直後（Closing）。位置保存のみ。重い切断は Closed 後に行う。</summary>
        public void OnMainWindowClosing(MainWindow window)
        {
            if (window != _mainWindow)
                return;

            try { window.SaveWindowBoundsFromRuntime(); } catch { }
        }

        /// <summary>アプリを完全終了する唯一の入口。</summary>
        /// <param name="closeWindow">false のとき窓は既に閉じ済み（×経路）。再 Close しない。</param>
        public void ExitApplication(bool closeWindow = true)
        {
            // 二重呼び出し防止（× Closed と --exit リスナー等）
            if (Interlocked.Exchange(ref _exitStarted, 1) != 0)
                return;

            _isExitingProcess = true;

            try
            {
                _listenerCts?.Cancel();
                _listenerCts?.Dispose();
            }
            catch { }
            _listenerCts = null;

            MainWindow? window = _mainWindow;
            _mainWindow = null;

            if (closeWindow && window != null)
            {
                try { window.Closed -= MainWindow_Closed; } catch { }
                try { window.Close(); } catch { }
            }

            // シリアル Close 待ちで UI を固めない（Dispose 内は非同期切断）
            try { _device.Dispose(); } catch { }

            try { SingleInstanceManager.Release(); } catch { }

            try { _app.Exit(); } catch { }
        }

        public void Dispose() => ExitApplication();

        private void MainWindow_Closed(object sender, WindowEventArgs e)
        {
            try { ((MainWindow)sender).Closed -= MainWindow_Closed; } catch { }

            if (ReferenceEquals(_mainWindow, sender))
                _mainWindow = null;
            else if (_mainWindow == null)
            {
                // 既に Exit 側で null 化済みでも続行可
            }

            // トレイなし：窓が閉じたら必ずプロセス終了（再 Close はしない）
            ExitApplication(closeWindow: false);
        }

        private void StartListeners()
        {
            var showEvent = SingleInstanceManager.InteractiveShowEvent;
            var exitEvent = SingleInstanceManager.ExitEvent;
            if (showEvent == null && exitEvent == null)
                return;

            _listenerCts = new CancellationTokenSource();
            var token = _listenerCts.Token;

            if (showEvent != null)
                Task.Run(() => ListenLoop(showEvent, token, () => ShowOrCreateMainWindow()), token);

            if (exitEvent != null)
                Task.Run(() => ListenLoop(exitEvent, token, () => _dispatcherQueue.TryEnqueue(() => ExitApplication())), token);
        }

        private static void ListenLoop(EventWaitHandle handle, CancellationToken token, Action action)
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    if (!handle.WaitOne(500))
                        continue;
                }
                catch (ObjectDisposedException)
                {
                    break;
                }

                if (token.IsCancellationRequested)
                    break;

                action();
            }
        }

        private static void BringWindowToForeground(Window window)
        {
            window.AppWindow.IsShownInSwitchers = true;
            window.Activate();
            IntPtr hwnd = WindowNative.GetWindowHandle(window);
            if (hwnd != IntPtr.Zero)
                SetForegroundWindow(hwnd);
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
