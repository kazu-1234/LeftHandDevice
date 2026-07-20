// v2.0.7
// StackPanel / Grid 向け縦ドラッグ並び替え（ブロック DnD はドラッグ中レイアウト再配置なし）
using System;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Windows.Foundation;

namespace LeftHandDevice
{
    /// <summary>
    /// ハンドルを掴んで項目を縦ドラッグ並び替えする。
    /// </summary>
    public static class DragReorderHelper
    {
        /// <summary>固定行高の Grid 用（ステップ行など）。</summary>
        public static void Attach(
            FrameworkElement handle,
            FrameworkElement item,
            Grid container,
            double itemHeight,
            Action onOrderChanged)
        {
            var translate = EnsureTranslate(item);
            bool isDragging = false;
            Point startPos = default;
            uint pointerId = 0;
            const double ThresholdRatio = 0.55;

            handle.PointerPressed += (s, e) =>
            {
                if (!e.GetCurrentPoint(handle).Properties.IsLeftButtonPressed)
                    return;
                isDragging = true;
                pointerId = e.Pointer.PointerId;
                translate.Y = 0;
                startPos = e.GetCurrentPoint(container).Position;
                handle.CapturePointer(e.Pointer);
                Canvas.SetZIndex(item, 100);
                item.Opacity = 0.88;
                e.Handled = true;
            };

            handle.PointerMoved += (s, e) =>
            {
                if (!isDragging || e.Pointer.PointerId != pointerId)
                    return;

                Point current = e.GetCurrentPoint(container).Position;
                double offsetY = current.Y - startPos.Y;
                translate.Y = offsetY;

                int count = container.Children.Count;
                if (count <= 1) { e.Handled = true; return; }

                double thr = itemHeight * ThresholdRatio;
                int currentIndex = Grid.GetRow(item);
                int direction = offsetY >= thr ? 1 : offsetY <= -thr ? -1 : 0;
                if (direction != 0)
                {
                    int targetIndex = Math.Clamp(currentIndex + direction, 0, count - 1);
                    if (targetIndex != currentIndex)
                    {
                        MoveGridRowByOne(container, currentIndex, targetIndex, item);
                        startPos = new Point(startPos.X, startPos.Y + direction * itemHeight);
                        translate.Y = current.Y - startPos.Y;
                    }
                }
                e.Handled = true;
            };

            void FinishGrid(bool cb)
            {
                if (!isDragging) return;
                isDragging = false;
                item.Opacity = 1;
                Canvas.SetZIndex(item, 0);
                translate.Y = 0;
                if (!cb) return;
                container.DispatcherQueue.TryEnqueue(DispatcherQueuePriority.Normal, () =>
                {
                    try { onOrderChanged(); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine(ex); }
                });
            }

            handle.PointerReleased += (s, e) =>
            {
                if (!isDragging || e.Pointer.PointerId != pointerId) return;
                try { handle.ReleasePointerCapture(e.Pointer); } catch { }
                FinishGrid(true);
                e.Handled = true;
            };
            handle.PointerCaptureLost += (s, e) => FinishGrid(true);
            handle.PointerCanceled += (s, e) =>
            {
                if (!isDragging || e.Pointer.PointerId != pointerId) return;
                try { handle.ReleasePointerCapture(e.Pointer); } catch { }
                FinishGrid(true);
                e.Handled = true;
            };
        }

        /// <summary>
        /// 可変高 StackPanel 用。ドラッグ中は Children を動かさず Transform のみで隙間を作り、
        /// ドロップ時に一度だけ並び替える（かくかく防止）。
        /// 画面端付近では親 ScrollViewer を自動スクロールする。
        /// </summary>
        public static void AttachToStackPanel(
            FrameworkElement handle,
            FrameworkElement item,
            StackPanel container,
            Action onOrderChanged)
        {
            var translate = EnsureTranslate(item);
            bool isDragging = false;
            Point pressPos = default;
            uint pointerId = 0;
            int startIndex = -1;
            int hoverIndex = -1;
            double lastContainerY = 0;
            double lastViewportY = 0;
            double autoScrollDir = 0;      // -1:上 / +1:下
            double autoScrollIntensity = 0; // 0..1

            ScrollViewer? scrollViewer = FindAncestorScrollViewer(container);
            DispatcherQueueTimer? autoScrollTimer = null;

            // 端ゾーンを広めに取り、端に近いほどはっきり加速する
            const double EdgeZonePx = 80;
            const double MinScrollPerTick = 2;
            const double MaxScrollPerTick = 48;

            void StopAutoScroll()
            {
                autoScrollDir = 0;
                autoScrollIntensity = 0;
                if (autoScrollTimer != null)
                {
                    autoScrollTimer.Stop();
                    autoScrollTimer.Tick -= AutoScrollTimer_Tick;
                    autoScrollTimer = null;
                }
            }

            void EnsureAutoScrollTimer()
            {
                if (autoScrollTimer != null || scrollViewer == null)
                    return;
                autoScrollTimer = container.DispatcherQueue.CreateTimer();
                autoScrollTimer.IsRepeating = true;
                autoScrollTimer.Interval = TimeSpan.FromMilliseconds(16);
                autoScrollTimer.Tick += AutoScrollTimer_Tick;
                autoScrollTimer.Start();
            }

            void UpdateDragFromContainerY(double containerY)
            {
                lastContainerY = containerY;
                translate.Y = containerY - pressPos.Y;

                int newHover = HitTestIndex(container, containerY);
                if (newHover != hoverIndex)
                {
                    hoverIndex = newHover;
                    ApplySiblingGapTransforms(container, item, startIndex, hoverIndex);
                }
            }

            /// <summary>
            /// ポインタの ScrollViewer ビューポート内 Y を取得する。
            /// GetCurrentPoint(sv) は環境によってずれることがあるため Transform で算出する。
            /// </summary>
            double GetViewportY(PointerRoutedEventArgs e)
            {
                if (scrollViewer == null)
                    return 0;

                UIElement? root = scrollViewer.XamlRoot?.Content as UIElement;
                if (root != null)
                {
                    try
                    {
                        Point pointerInRoot = e.GetCurrentPoint(root).Position;
                        Point viewerOrigin = scrollViewer.TransformToVisual(root)
                            .TransformPoint(new Point(0, 0));
                        return pointerInRoot.Y - viewerOrigin.Y;
                    }
                    catch
                    {
                        // fall through
                    }
                }

                return e.GetCurrentPoint(scrollViewer).Position.Y;
            }

            /// <summary>PageScrollHost が Disabled にしていても、内容が溢れていれば一時的に有効化する。</summary>
            void EnsureVerticalScrollEnabled()
            {
                if (scrollViewer == null)
                    return;

                double viewportH = scrollViewer.ViewportHeight;
                if (viewportH < 1)
                    viewportH = scrollViewer.ActualHeight;
                if (viewportH < 1)
                    return;

                if (scrollViewer.ExtentHeight <= viewportH + 0.5)
                    return;

                // PageScrollHost 経由なら専用 API を使う
                if (FindAncestorPageScrollHost(container) is Views.PageScrollHost host)
                {
                    host.EnsureVerticalScrollEnabled();
                    return;
                }

                if (scrollViewer.VerticalScrollMode == ScrollMode.Disabled)
                {
                    scrollViewer.VerticalScrollMode = ScrollMode.Enabled;
                    scrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
                }
            }

            /// <summary>ビューポート座標 Y から端スクロール方向・強度を更新する。</summary>
            void ApplyAutoScrollFromViewportY(double viewportY)
            {
                if (scrollViewer == null)
                    return;

                double viewportH = scrollViewer.ViewportHeight;
                if (viewportH < 1)
                    viewportH = scrollViewer.ActualHeight;
                if (viewportH < 1)
                {
                    autoScrollDir = 0;
                    autoScrollIntensity = 0;
                    return;
                }

                // ビューポート相対のみ使う（VerticalOffset 補正はしない）
                double y = Math.Clamp(viewportY, 0, viewportH);
                lastViewportY = y;

                if (y < EdgeZonePx)
                {
                    double linear = Math.Clamp((EdgeZonePx - y) / EdgeZonePx, 0, 1);
                    autoScrollDir = -1;
                    autoScrollIntensity = linear * linear;
                    EnsureAutoScrollTimer();
                }
                else if (y > viewportH - EdgeZonePx)
                {
                    double linear = Math.Clamp((y - (viewportH - EdgeZonePx)) / EdgeZonePx, 0, 1);
                    autoScrollDir = 1;
                    autoScrollIntensity = linear * linear;
                    EnsureAutoScrollTimer();
                }
                else
                {
                    autoScrollDir = 0;
                    autoScrollIntensity = 0;
                }
            }

            void AutoScrollTimer_Tick(DispatcherQueueTimer sender, object args)
            {
                if (!isDragging || scrollViewer == null)
                    return;

                ApplyAutoScrollFromViewportY(lastViewportY);
                if (autoScrollDir == 0)
                    return;

                EnsureVerticalScrollEnabled();

                double speed = MinScrollPerTick
                    + (MaxScrollPerTick - MinScrollPerTick) * autoScrollIntensity;
                double max = Math.Max(0, scrollViewer.ExtentHeight - scrollViewer.ViewportHeight);
                if (max < 0.5)
                    max = scrollViewer.ScrollableHeight;

                double target = Math.Clamp(
                    scrollViewer.VerticalOffset + autoScrollDir * speed,
                    0,
                    max);
                double actualDelta = target - scrollViewer.VerticalOffset;
                if (Math.Abs(actualDelta) < 0.1)
                    return;

                scrollViewer.ChangeView(null, target, null, disableAnimation: true);

                // ポインタ固定のまま内容が動くので、コンテナ座標の Y を補正
                UpdateDragFromContainerY(lastContainerY + actualDelta);
            }

            handle.PointerPressed += (s, e) =>
            {
                if (!e.GetCurrentPoint(handle).Properties.IsLeftButtonPressed)
                    return;

                startIndex = container.Children.IndexOf(item);
                if (startIndex < 0) return;

                scrollViewer ??= FindAncestorScrollViewer(container);
                isDragging = true;
                hoverIndex = startIndex;
                pointerId = e.Pointer.PointerId;
                translate.Y = 0;
                pressPos = e.GetCurrentPoint(container).Position;
                lastContainerY = pressPos.Y;
                autoScrollDir = 0;
                autoScrollIntensity = 0;
                if (scrollViewer != null)
                {
                    EnsureVerticalScrollEnabled();
                    ApplyAutoScrollFromViewportY(GetViewportY(e));
                }
                handle.CapturePointer(e.Pointer);
                Canvas.SetZIndex(item, 200);
                item.Opacity = 0.92;
                e.Handled = true;
            };

            handle.PointerMoved += (s, e) =>
            {
                if (!isDragging || e.Pointer.PointerId != pointerId)
                    return;

                Point current = e.GetCurrentPoint(container).Position;
                UpdateDragFromContainerY(current.Y);

                if (scrollViewer != null)
                    ApplyAutoScrollFromViewportY(GetViewportY(e));

                e.Handled = true;
            };

            void FinishPanel(bool commit)
            {
                if (!isDragging) return;
                isDragging = false;
                StopAutoScroll();

                int from = startIndex;
                int to = hoverIndex;
                startIndex = -1;
                hoverIndex = -1;

                ClearAllTransforms(container);
                item.Opacity = 1.0;
                Canvas.SetZIndex(item, 0);
                translate.Y = 0;

                if (commit && from >= 0 && to >= 0 && from != to
                    && from < container.Children.Count)
                {
                    UIElement moved = container.Children[from];
                    container.Children.RemoveAt(from);
                    int insertAt = Math.Clamp(to, 0, container.Children.Count);
                    container.Children.Insert(insertAt, moved);
                }

                if (!commit) return;

                container.DispatcherQueue.TryEnqueue(DispatcherQueuePriority.Normal, () =>
                {
                    try { onOrderChanged(); }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine(ex); }
                });
            }

            handle.PointerReleased += (s, e) =>
            {
                if (!isDragging || e.Pointer.PointerId != pointerId) return;
                try { handle.ReleasePointerCapture(e.Pointer); } catch { }
                FinishPanel(true);
                e.Handled = true;
            };
            handle.PointerCaptureLost += (s, e) =>
            {
                if (!isDragging) return;
                FinishPanel(true);
            };
            handle.PointerCanceled += (s, e) =>
            {
                if (!isDragging || e.Pointer.PointerId != pointerId) return;
                try { handle.ReleasePointerCapture(e.Pointer); } catch { }
                FinishPanel(true);
                e.Handled = true;
            };
        }

        private static ScrollViewer? FindAncestorScrollViewer(DependencyObject start)
        {
            DependencyObject? current = VisualTreeHelper.GetParent(start);
            while (current != null)
            {
                if (current is ScrollViewer sv)
                    return sv;
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }

        private static Views.PageScrollHost? FindAncestorPageScrollHost(DependencyObject start)
        {
            DependencyObject? current = VisualTreeHelper.GetParent(start);
            while (current != null)
            {
                if (current is Views.PageScrollHost host)
                    return host;
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }

        private static int HitTestIndex(StackPanel container, double pointerY)
        {
            int count = container.Children.Count;
            if (count <= 1) return 0;

            double acc = 0;
            for (int i = 0; i < count; i++)
            {
                if (container.Children[i] is not FrameworkElement fe) continue;
                double h = fe.ActualHeight;
                if (h < 1) h = 40;
                double mid = acc + h * 0.5;
                if (pointerY < mid)
                    return i;
                acc += h + container.Spacing;
            }
            return count - 1;
        }

        private static void ApplySiblingGapTransforms(
            StackPanel container,
            FrameworkElement dragged,
            int startIndex,
            int hoverIndex)
        {
            double dragH = dragged.ActualHeight;
            if (dragH < 1) dragH = 80;
            dragH += container.Spacing;

            for (int i = 0; i < container.Children.Count; i++)
            {
                if (container.Children[i] is not FrameworkElement fe)
                    continue;
                if (ReferenceEquals(fe, dragged))
                    continue;

                double shift = 0;
                if (startIndex < hoverIndex)
                {
                    if (i > startIndex && i <= hoverIndex)
                        shift = -dragH;
                }
                else if (startIndex > hoverIndex)
                {
                    if (i >= hoverIndex && i < startIndex)
                        shift = dragH;
                }

                AnimateTranslateY(fe, shift);
            }
        }

        private static void ClearAllTransforms(StackPanel container)
        {
            foreach (UIElement child in container.Children)
            {
                if (child is not FrameworkElement fe) continue;
                var t = EnsureTranslate(fe);
                // 進行中アニメを打ち消すため直接代入
                t.Y = 0;
            }
        }

        private static void AnimateTranslateY(FrameworkElement fe, double to)
        {
            var t = EnsureTranslate(fe);
            // ドラッグ中の隙間移動は即時反映（Children 再配置がないので十分なめらか）
            t.Y = to;
        }

        private static TranslateTransform EnsureTranslate(FrameworkElement fe)
        {
            if (fe.RenderTransform is TranslateTransform existing)
                return existing;
            var t = new TranslateTransform();
            fe.RenderTransform = t;
            return t;
        }

        private static void MoveGridRowByOne(
            Grid container,
            int fromRow,
            int toRow,
            FrameworkElement dragged)
        {
            if (fromRow == toRow) return;

            foreach (UIElement child in container.Children)
            {
                if (child is not FrameworkElement fe) continue;
                int row = Grid.GetRow(fe);
                int newRow = row;

                if (fromRow < toRow)
                {
                    if (fe == dragged) newRow = toRow;
                    else if (row > fromRow && row <= toRow) newRow = row - 1;
                }
                else
                {
                    if (fe == dragged) newRow = toRow;
                    else if (row >= toRow && row < fromRow) newRow = row + 1;
                }

                if (newRow != row)
                    Grid.SetRow(fe, newRow);
            }
        }

        public static object?[] GetOrderedTags(Grid container)
        {
            int count = container.Children.Count;
            var result = new object?[count];
            foreach (UIElement child in container.Children)
            {
                if (child is not FrameworkElement fe) continue;
                int row = Grid.GetRow(fe);
                if (row >= 0 && row < count)
                    result[row] = fe.Tag;
            }
            return result;
        }

        public static object?[] GetOrderedTags(StackPanel container)
        {
            var result = new object?[container.Children.Count];
            for (int i = 0; i < container.Children.Count; i++)
            {
                if (container.Children[i] is FrameworkElement fe)
                    result[i] = fe.Tag;
            }
            return result;
        }
    }
}
