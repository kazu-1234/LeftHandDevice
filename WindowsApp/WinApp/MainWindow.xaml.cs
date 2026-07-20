using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media.Animation;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.IO;
using System.Linq;
using LeftHandDevice.Views;

namespace LeftHandDevice
{
    public sealed partial class MainWindow : Window
    {
        private const int DefaultClientWidth = 960;
        private const int DefaultClientHeight = 700;
        private const double MinimumWindowWidth = 870;
        private const double MinimumWindowHeight = 600;

        private readonly AppRuntime _runtime;
        private readonly AppState _appState;
        private bool _windowBoundsReady;
        private string _currentPageTag = "Home";
        private TitleBarThemeHelper? _titleBarThemeHelper;
        private string _baseTitle = "LeftHandDevice";

        public MainWindow(AppRuntime runtime)
        {
            _runtime = runtime;
            _appState = runtime.AppState;

            InitializeComponent();
            _baseTitle = Strings.Get("AppName");
            Title = _baseTitle;
            AppTitleBar.Title = _baseTitle;
            ApplyWindowIcon();

            ThemeService.AttachRoot(RootGrid);
            _titleBarThemeHelper = new TitleBarThemeHelper(this, RootGrid, AppTitleBar);

            AppWindow.Closing += AppWindow_Closing;
            AppWindow.Changed += AppWindow_Changed;
            ContentFrame.NavigationFailed += ContentFrame_NavigationFailed;

            _appState.Device.VolumeHintChanged += OnVolumeHintChanged;
            _appState.Device.ContinuousWarningRequested += OnContinuousWarning;

            ConfigureMinimumWindowSize();
            Activated += MainWindow_Activated;
        }

        public void PrepareAndActivate(string? initialPageTag = null)
        {
            // 先にウィンドウを表示してからページ遷移する
            //（HomePage の自動 COM 接続などが UI をブロックしても、窓が出ない状態を防ぐ）
            AppWindow.IsShownInSwitchers = true;
            RestoreWindowBounds();
            Activate();
            try { AppWindow.Show(true); } catch { try { AppWindow.Show(); } catch { } }

            IntPtr hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            if (hwnd != IntPtr.Zero)
            {
                ShowWindow(hwnd, SwRestore);
                ShowWindow(hwnd, SwShow);
                SetForegroundWindow(hwnd);
            }

            string tag = string.IsNullOrEmpty(initialPageTag) ? GetDefaultPageTag() : initialPageTag;
            NavigateToPage(tag, force: true, suppressTransition: true);

            WindowPlacementHelper.ApplyMaximizedIfNeeded(this, _runtime.Settings);
            _windowBoundsReady = true;
            Activate();
        }

        private const int SwShow = 5;
        private const int SwRestore = 9;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        public void NavigateToPageTag(string tag) => NavigateToPage(tag, force: false, suppressTransition: true);

        internal void SaveWindowBoundsFromRuntime()
        {
            try { WindowPlacementHelper.Save(this, _runtime.Settings); }
            catch { }
        }

        private static string GetDefaultPageTag() => "Home";

        private void ApplyWindowIcon()
        {
            string iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico");
            if (!File.Exists(iconPath))
                return;

            try { AppWindow.SetIcon(iconPath); }
            catch { }
        }

        private void ContentFrame_NavigationFailed(object sender, NavigationFailedEventArgs e)
        {
            Title = $"{Strings.Get("AppName")} - {e.Exception?.Message}";
        }

        private void ConfigureMinimumWindowSize()
        {
            if (AppWindow.Presenter is not OverlappedPresenter presenter)
                return;

            presenter.IsResizable = true;
            double scaleFactor = RootGrid.XamlRoot?.RasterizationScale ?? 1.0;
            presenter.PreferredMinimumWidth = (int)(MinimumWindowWidth * scaleFactor);
            presenter.PreferredMinimumHeight = (int)(MinimumWindowHeight * scaleFactor);
            presenter.PreferredMaximumWidth = 10000;
            presenter.PreferredMaximumHeight = 10000;
        }

        private void AppWindow_Changed(AppWindow sender, AppWindowChangedEventArgs args)
        {
            if (_runtime.IsExitingProcess || !_windowBoundsReady)
                return;

            if (args.DidSizeChange || args.DidPositionChange)
                SaveWindowBounds();
        }

        private void SaveWindowBounds()
        {
            if (!_windowBoundsReady || _runtime.IsExitingProcess)
                return;

            try { WindowPlacementHelper.Save(this, _runtime.Settings); }
            catch { }
        }

        private void RestoreWindowBounds()
        {
            WindowPlacementHelper.Restore(this, _runtime.Settings, DefaultClientWidth, DefaultClientHeight);
        }

        private void AppWindow_Closing(AppWindow sender, AppWindowClosingEventArgs args)
        {
            // 既に Exit 進行中なら追加処理しない（二重 Close 経路の干渉を防ぐ）
            if (_runtime.IsExitingProcess)
            {
                args.Cancel = false;
                return;
            }

            // 閉じる直前の位置を保存してからレース防止フラグを落とす
            try { WindowPlacementHelper.Save(this, _runtime.Settings); } catch { }
            _windowBoundsReady = false;

            try { _runtime.OnMainWindowClosing(this); } catch { }

            try
            {
                AppWindow.Closing -= AppWindow_Closing;
                AppWindow.Changed -= AppWindow_Changed;
                _appState.Device.VolumeHintChanged -= OnVolumeHintChanged;
                _appState.Device.ContinuousWarningRequested -= OnContinuousWarning;
                Activated -= MainWindow_Activated;
            }
            catch { }

            // ページ内フック等を先に外す（終了フリーズ防止）
            try
            {
                if (ContentFrame.Content is HomePage home)
                    home.PrepareForClose();
            }
            catch { }

            // 窓のクローズは続行。プロセス終了は Closed → ExitApplication(closeWindow:false)
            args.Cancel = false;
        }

        private void MainWindow_Activated(object sender, WindowActivatedEventArgs args)
        {
            if (args.WindowActivationState == WindowActivationState.Deactivated)
                _appState.Device.RequestContinuousWarning();
        }

        private void OnVolumeHintChanged(string? hint)
        {
            Title = string.IsNullOrEmpty(hint) ? _baseTitle : $"{_baseTitle} — {hint}";
            AppTitleBar.Title = Title;
        }

        private void OnContinuousWarning()
        {
            // HomePage 側でオーバーレイ表示するため、ナビ状態を維持しつつイベントは Device から配信済み
        }

        private void NavView_ItemInvoked(NavigationView sender, NavigationViewItemInvokedEventArgs args)
        {
            if (args.InvokedItemContainer is NavigationViewItem item && item.Tag is string tag)
                NavigateToPage(tag);
        }

        private void NavigateToPage(string tag, bool force = false, bool suppressTransition = false)
        {
            if (!force && _currentPageTag == tag && ContentFrame.CurrentSourcePageType != null)
            {
                UpdateNavSelection(tag);
                return;
            }

            _currentPageTag = tag;
            Type pageType = tag switch
            {
                "Info" => typeof(InfoPage),
                "Settings" => typeof(SettingsPage),
                _ => typeof(HomePage)
            };

            if (force || ContentFrame.CurrentSourcePageType != pageType)
            {
                if (suppressTransition)
                    ContentFrame.Navigate(pageType, _appState, new SuppressNavigationTransitionInfo());
                else
                    ContentFrame.Navigate(pageType, _appState);
            }

            UpdateNavSelection(tag);
        }

        private void UpdateNavSelection(string tag)
        {
            NavigationViewItem? match = null;
            foreach (var item in NavView.MenuItems.OfType<NavigationViewItem>())
            {
                if (item.Tag as string == tag)
                    match = item;
            }

            foreach (var item in NavView.FooterMenuItems.OfType<NavigationViewItem>())
            {
                if (item.Tag as string == tag)
                    match = item;
            }

            if (match != null)
                NavView.SelectedItem = match;
            else if (NavItemHome != null)
                NavView.SelectedItem = NavItemHome;
        }
    }
}
