// v2.0.18
// マウス一括登録中のクリック位置に番号マーカーを画面上へ表示する。
// v1.14.0 の WPF 実装と同方式: System.Windows.Window + AllowsTransparency で透明背景。
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using WpfWindow = System.Windows.Window;
using WpfWindowStyle = System.Windows.WindowStyle;
using WpfResizeMode = System.Windows.ResizeMode;
using SWM = System.Windows.Media;

namespace LeftHandDevice
{
    internal sealed class ClickMarkerOverlay : IDisposable
    {
        private const int MarkerSizeDip = 36;

        private static readonly string[] CircledNumbers =
        {
            "❶", "❷", "❸", "❹", "❺", "❻", "❼", "❽", "❾", "❿"
        };

        // ブラシ・角丸は使い回す（毎回 new しない）
        private static readonly SWM.SolidColorBrush _markerBrush;
        private static readonly System.Windows.CornerRadius _markerCornerRadius
            = new(MarkerSizeDip / 2.0);

        // DPI はプロセス起動時に 1 回取得してキャッシュ
        private static readonly double _dpiScaleX;
        private static readonly double _dpiScaleY;

        static ClickMarkerOverlay()
        {
            _markerBrush = new SWM.SolidColorBrush(
                SWM.Color.FromArgb(220, 255, 80, 80));
            _markerBrush.Freeze(); // Freeze でスレッド安全＋描画高速化

            double dx = 1.0, dy = 1.0;
            GetDpiScale(ref dx, ref dy);
            _dpiScaleX = dx;
            _dpiScaleY = dy;
        }

        private readonly List<WpfWindow> _markers = new();
        private Thread? _wpfThread;
        private System.Windows.Threading.Dispatcher? _wpfDispatcher;
        private readonly ManualResetEventSlim _dispatcherReady = new(false);

        /// <summary>
        /// WPF Dispatcher スレッドを起動する（初回のみ）。
        /// WinUI プロセス内で WPF ウィンドウを表示するには、
        /// 専用の STA スレッドで Dispatcher を走らせる必要がある。
        /// </summary>
        private void EnsureWpfThread()
        {
            if (_wpfDispatcher != null) return;

            _wpfThread = new Thread(() =>
            {
                _wpfDispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher;
                _dispatcherReady.Set();
                System.Windows.Threading.Dispatcher.Run();
            });
            _wpfThread.SetApartmentState(ApartmentState.STA);
            _wpfThread.IsBackground = true;
            _wpfThread.Start();
            _dispatcherReady.Wait();
        }

        public void Show(int screenX, int screenY, int number)
        {
            EnsureWpfThread();
            if (_wpfDispatcher == null) return;

            string label = number <= 10 ? CircledNumbers[number - 1] : number.ToString();

            // DPI 補正: フック座標は物理ピクセル、WPF は DIP（キャッシュ済み）
            double wpfX = screenX / _dpiScaleX;
            double wpfY = screenY / _dpiScaleY;

            // 非同期で WPF スレッドへ投げる（呼び出し元をブロックしない）
            _wpfDispatcher.InvokeAsync(() =>
            {
                var markerWindow = new WpfWindow
                {
                    WindowStyle = WpfWindowStyle.None,
                    AllowsTransparency = true,
                    Background = SWM.Brushes.Transparent,
                    Topmost = true,
                    ShowInTaskbar = false,
                    IsHitTestVisible = false,
                    Width = MarkerSizeDip,
                    Height = MarkerSizeDip,
                    Left = wpfX - (MarkerSizeDip / 2.0),
                    Top = wpfY - (MarkerSizeDip / 2.0),
                    ResizeMode = WpfResizeMode.NoResize
                };

                var border = new System.Windows.Controls.Border
                {
                    Background = _markerBrush,
                    CornerRadius = _markerCornerRadius,
                    Width = MarkerSizeDip,
                    Height = MarkerSizeDip
                };

                var text = new System.Windows.Controls.TextBlock
                {
                    Text = label,
                    Foreground = SWM.Brushes.White,
                    FontSize = 18,
                    FontWeight = System.Windows.FontWeights.Bold,
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    TextAlignment = System.Windows.TextAlignment.Center
                };

                border.Child = text;
                markerWindow.Content = border;
                markerWindow.Show();
                _markers.Add(markerWindow);
            });
        }

        public void Clear()
        {
            if (_wpfDispatcher == null) return;

            // Clear は Stop() から呼ばれるため同期で待つ（マーカーが確実に消えるように）
            // ただしタイムアウト付きでデッドロックを防止
            _wpfDispatcher.Invoke(() =>
            {
                foreach (WpfWindow marker in _markers)
                {
                    try { marker.Close(); } catch { }
                }
                _markers.Clear();
            }, System.Windows.Threading.DispatcherPriority.Send,
            System.Threading.CancellationToken.None,
            TimeSpan.FromMilliseconds(2000));
        }

        public void Dispose()
        {
            Clear();
            if (_wpfDispatcher != null)
            {
                _wpfDispatcher.InvokeShutdown();
                _wpfDispatcher = null;
            }
        }

        private static void GetDpiScale(ref double dpiX, ref double dpiY)
        {
            IntPtr hdc = GetDC(IntPtr.Zero);
            if (hdc == IntPtr.Zero) return;
            dpiX = GetDeviceCaps(hdc, 88) / 96.0; // LOGPIXELSX
            dpiY = GetDeviceCaps(hdc, 90) / 96.0; // LOGPIXELSY
            ReleaseDC(IntPtr.Zero, hdc);
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern int GetDeviceCaps(IntPtr hdc, int index);
    }
}
