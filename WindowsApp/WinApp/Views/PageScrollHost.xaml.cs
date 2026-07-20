using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Windows.Foundation;

namespace LeftHandDevice.Views
{
    public sealed partial class PageScrollHost : ContentControl
    {
        private ScrollViewer? _scrollViewer;
        private FrameworkElement? _contentRoot;
        private bool _scrollEnabled;
        private bool _updateScheduled;

        public PageScrollHost()
        {
            InitializeComponent();
            Loaded += OnLoaded;
            SizeChanged += (_, __) => ScheduleUpdateScrollability();
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            _scrollViewer = GetTemplateChild("PART_ScrollViewer") as ScrollViewer;
            if (_scrollViewer != null)
            {
                _scrollViewer.SizeChanged += (_, __) => ScheduleUpdateScrollability();
                _scrollViewer.PointerWheelChanged += ScrollViewer_PointerWheelChanged;
            }

            WatchContentRoot(Content as FrameworkElement);
            ScheduleUpdateScrollability();
        }

        protected override void OnContentChanged(object oldContent, object newContent)
        {
            base.OnContentChanged(oldContent, newContent);

            if (oldContent is FrameworkElement oldRoot)
                UnwatchContentRoot(oldRoot);

            WatchContentRoot(newContent as FrameworkElement);
            ScheduleUpdateScrollability();
        }

        private void WatchContentRoot(FrameworkElement? root)
        {
            if (root == null)
                return;

            _contentRoot = root;
            root.SizeChanged += ContentRoot_SizeChanged;
            root.Loaded += ContentRoot_Loaded;
        }

        private void UnwatchContentRoot(FrameworkElement root)
        {
            root.SizeChanged -= ContentRoot_SizeChanged;
            root.Loaded -= ContentRoot_Loaded;

            if (ReferenceEquals(_contentRoot, root))
                _contentRoot = null;
        }

        private void ContentRoot_Loaded(object sender, RoutedEventArgs e)
        {
            ScheduleUpdateScrollability();
        }

        private void ContentRoot_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            ScheduleUpdateScrollability();
        }

        private void ScrollViewer_PointerWheelChanged(object sender, PointerRoutedEventArgs e)
        {
            if (!_scrollEnabled)
                e.Handled = true;
        }

        /// <summary>パターン描画後など、明示的にスクロール要否を再計算する。</summary>
        public void InvalidateScrollability() => ScheduleUpdateScrollability();

        /// <summary>プログラムからのスクロール（DnD 自動スクロール等）を許可する。</summary>
        public void EnsureVerticalScrollEnabled()
        {
            if (_scrollViewer == null)
                _scrollViewer = GetTemplateChild("PART_ScrollViewer") as ScrollViewer;
            if (_scrollViewer == null)
                return;

            // 要否を再測定してから有効化
            if (ComputeNeedsScroll())
                ApplyScrollState(true);
        }

        /// <summary>テンプレート内 ScrollViewer（無ければ null）。</summary>
        public ScrollViewer? HostScrollViewer
        {
            get
            {
                if (_scrollViewer == null)
                    _scrollViewer = GetTemplateChild("PART_ScrollViewer") as ScrollViewer;
                return _scrollViewer;
            }
        }

        private void ScheduleUpdateScrollability()
        {
            if (_updateScheduled)
                return;

            _updateScheduled = true;
            DispatcherQueue.TryEnqueue(DispatcherQueuePriority.Low, () =>
            {
                _updateScheduled = false;
                UpdateScrollability();
            });
        }

        private void UpdateScrollability()
        {
            if (_scrollViewer == null)
                return;

            bool needsScroll = ComputeNeedsScroll();
            ApplyScrollState(needsScroll);

            if (!needsScroll && _scrollViewer.VerticalOffset > 0)
                _scrollViewer.ChangeView(null, 0, null, disableAnimation: true);
        }

        /// <summary>
        /// スクロール無効時は子がビューポート高さにクリップされ ActualHeight が当てにならない。
        /// 無限高さで Measure し、内容の希望高さとビューポートを比較する。
        /// </summary>
        private bool ComputeNeedsScroll()
        {
            if (_scrollViewer == null)
                return false;

            double viewportHeight = _scrollViewer.ViewportHeight;
            if (viewportHeight <= 0)
                viewportHeight = _scrollViewer.ActualHeight;
            if (viewportHeight <= 0)
                return false;

            double contentHeight = MeasureContentHeight();
            if (contentHeight <= 0)
                contentHeight = _scrollViewer.ExtentHeight;

            return contentHeight > viewportHeight + 0.5;
        }

        private double MeasureContentHeight()
        {
            if (_contentRoot == null)
                return 0;

            double width = _scrollViewer?.ViewportWidth ?? 0;
            if (width <= 0)
                width = _scrollViewer?.ActualWidth ?? 0;
            if (width <= 0)
                width = _contentRoot.ActualWidth;
            if (width <= 0)
                width = ActualWidth;

            // 横幅は実幅、縦は無限で内容の intrinsic 高さを取る
            _contentRoot.Measure(new Size(width > 0 ? width : double.PositiveInfinity, double.PositiveInfinity));
            double height = _contentRoot.DesiredSize.Height;

            // Grid が Stretch で DesiredSize をビューポートに合わせる場合に備え、子も見る
            if (_contentRoot is Panel panel)
            {
                foreach (UIElement child in panel.Children)
                {
                    if (child is FrameworkElement fe)
                    {
                        fe.Measure(new Size(width > 0 ? width : double.PositiveInfinity, double.PositiveInfinity));
                        height = Math.Max(height, fe.DesiredSize.Height);
                    }
                }
            }

            return height;
        }

        private void ApplyScrollState(bool enabled)
        {
            if (_scrollViewer == null)
                return;

            _scrollEnabled = enabled;
            _scrollViewer.VerticalScrollMode = enabled ? ScrollMode.Auto : ScrollMode.Disabled;
            _scrollViewer.VerticalScrollBarVisibility = enabled
                ? ScrollBarVisibility.Auto
                : ScrollBarVisibility.Disabled;
        }
    }
}
