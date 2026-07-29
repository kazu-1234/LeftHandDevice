// HomePage.xaml.cs
// パターン編集・COM 接続・マウス座標キャプチャ UI
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Windows.UI.Core;
using VirtualKey = Windows.System.VirtualKey;

namespace LeftHandDevice.Views
{
    public sealed partial class HomePage : Page
    {
        private AppState? _state;
        private DeviceService? Device => _state?.Device;

        private bool _isRenderingPatterns;
        private bool _hasTriedAutoConnect;
        private bool _subscribed;

        private Microsoft.UI.Dispatching.DispatcherQueueTimer? _warningHideTimer;
        private MouseCaptureHelper? _mouseCapture;

        public HomePage()
        {
            InitializeComponent();

            _warningHideTimer = DispatcherQueue.CreateTimer();
            _warningHideTimer.IsRepeating = false;
            _warningHideTimer.Interval = TimeSpan.FromSeconds(3);
            _warningHideTimer.Tick += WarningHideTimer_Tick;
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            base.OnNavigatedTo(e);

            if (e.Parameter is not AppState state)
                return;

            // 再遷移時はいったん購読解除してから付け直す
            UnsubscribeDeviceEvents();
            _state = state;
            SubscribeDeviceEvents();

            UpdateConnectionUi();
            UpdateActiveButtonsText();
            _ = LoadComPortsAsync(scheduleAutoConnect: true);
            RenderAllPatterns();

            // 起動後の初回表示時のみアップデートを確認
            if (!_hasCheckedUpdateAtStartup)
            {
                _hasCheckedUpdateAtStartup = true;
                _ = CheckUpdateAtStartupAsync();
            }

            ComPortComboBox.DropDownOpened -= ComPortComboBox_DropDownOpened;
            ComPortComboBox.DropDownOpened += ComPortComboBox_DropDownOpened;
        }

        private static bool _hasCheckedUpdateAtStartup;

        private void ComPortComboBox_DropDownOpened(object? sender, object e)
        {
            // 開くたびに GetPortNames を UI で叩かない（フリーズ防止）
            if (Device == null || Device.IsConnected)
                return;
            _ = LoadComPortsAsync(scheduleAutoConnect: false);
        }

        private async System.Threading.Tasks.Task CheckUpdateAtStartupAsync()
        {
            try
            {
                await System.Threading.Tasks.Task.Delay(2000);
                if (XamlRoot == null) return;

                UpdateCheckResult result = await UpdateChecker.CheckForUpdateAsync();
                if (result.Status != UpdateCheckStatus.UpdateAvailable
                    || string.IsNullOrWhiteSpace(result.ReleasePageUrl))
                    return;

                var dialog = new ContentDialog
                {
                    Title = Strings.Get("Update_DialogTitle"),
                    Content = Strings.Format("Update_DialogContent", result.Message),
                    PrimaryButtonText = Strings.Get("Update_DialogOpen"),
                    CloseButtonText = "後で",
                    XamlRoot = XamlRoot
                };

                if (await dialog.ShowAsync() == ContentDialogResult.Primary)
                    await Windows.System.Launcher.LaunchUriAsync(new Uri(result.ReleasePageUrl));
            }
            catch { }
        }

        protected override void OnNavigatedFrom(NavigationEventArgs e)
        {
            base.OnNavigatedFrom(e);
            PrepareForClose();
        }

        /// <summary>ウィンドウ終了前にフック／購読を解放する。</summary>
        public void PrepareForClose()
        {
            try { ComPortComboBox.DropDownOpened -= ComPortComboBox_DropDownOpened; } catch { }
            try { _mouseCapture?.Stop(); } catch { }
            try { _mouseCapture?.Dispose(); } catch { }
            _mouseCapture = null;
            UnsubscribeDeviceEvents();
            if (_warningHideTimer != null)
            {
                _warningHideTimer.Stop();
                _warningHideTimer.Tick -= WarningHideTimer_Tick;
            }
        }

        private void SubscribeDeviceEvents()
        {
            if (Device == null || _subscribed) return;
            Device.PatternsChanged += Device_PatternsChanged;
            Device.ConnectionChanged += Device_ConnectionChanged;
            Device.ContinuousWarningRequested += Device_ContinuousWarningRequested;
            _subscribed = true;
        }

        private void UnsubscribeDeviceEvents()
        {
            if (Device == null || !_subscribed) return;
            Device.PatternsChanged -= Device_PatternsChanged;
            Device.ConnectionChanged -= Device_ConnectionChanged;
            Device.ContinuousWarningRequested -= Device_ContinuousWarningRequested;
            _subscribed = false;
        }

        private void Device_PatternsChanged()
        {
            DispatcherQueue.TryEnqueue(() =>
            {
                UpdateActiveButtonsText();
                RenderAllPatterns();
            });
        }

        private void Device_ConnectionChanged()
        {
            DispatcherQueue.TryEnqueue(UpdateConnectionUi);
        }

        private void Device_ContinuousWarningRequested()
        {
            DispatcherQueue.TryEnqueue(ShowContinuousWarning);
        }

        private void WarningHideTimer_Tick(Microsoft.UI.Dispatching.DispatcherQueueTimer sender, object args)
        {
            sender.Stop();
            ContinuousWarningOverlay.Visibility = Visibility.Collapsed;
        }

        private void ShowContinuousWarning()
        {
            ContinuousWarningOverlay.Visibility = Visibility.Visible;
            ContinuousWarningOverlay.Opacity = 1;
            _warningHideTimer?.Stop();
            _warningHideTimer?.Start();
        }

        private void UpdateActiveButtonsText()
        {
            if (Device == null) return;
            ActiveButtonsText.Text = $"有効ボタン数: {Device.ActiveButtonCount}";
        }

        private void UpdateConnectionUi()
        {
            if (Device == null) return;
            bool connected = Device.IsConnected;
            ConnectButton.Content = connected ? "切断する" : "接続する";
            ComPortComboBox.IsEnabled = !connected && !_isSerialBusy;
            ConnectButton.IsEnabled = !_isSerialBusy;
        }

        private bool _isSerialBusy;
        private bool _comPortsLoading;

        /// <summary>COM ポート一覧を BG で取得して ComboBox に反映する。</summary>
        private async System.Threading.Tasks.Task LoadComPortsAsync(bool scheduleAutoConnect)
        {
            if (Device == null || _comPortsLoading)
                return;

            _comPortsLoading = true;
            string? previousSelection = ComPortComboBox.SelectedItem as string;
            string? savedPort = Device.LoadSavedComPort();

            string[] ports;
            try
            {
                ports = await System.Threading.Tasks.Task.Run(() => Device.GetAvailablePorts());
            }
            catch
            {
                ports = Array.Empty<string>();
            }

            if (Device == null)
            {
                _comPortsLoading = false;
                return;
            }

            // ドロップダウン操作中でも選択を壊しすぎないよう、内容が変わったときだけ差し替え
            var current = new System.Collections.Generic.HashSet<string>(
                ComPortComboBox.Items.OfType<string>());
            bool same = ports.Length == current.Count && ports.All(current.Contains);
            if (!same)
            {
                ComPortComboBox.Items.Clear();
                foreach (string port in ports)
                    ComPortComboBox.Items.Add(port);
            }

            if (ComPortComboBox.Items.Count > 0)
            {
                if (!string.IsNullOrEmpty(savedPort)
                    && ComPortComboBox.Items.Contains(savedPort))
                {
                    ComPortComboBox.SelectedItem = savedPort;
                }
                else if (!string.IsNullOrEmpty(previousSelection)
                         && ComPortComboBox.Items.Contains(previousSelection))
                {
                    ComPortComboBox.SelectedItem = previousSelection;
                }
                else if (ComPortComboBox.SelectedItem == null)
                {
                    ComPortComboBox.SelectedIndex = 0;
                }
            }

            UpdateConnectionUi();
            _comPortsLoading = false;

            if (scheduleAutoConnect)
                ScheduleDeferredAutoConnect();
        }

        private void ScheduleDeferredAutoConnect()
        {
            if (_hasTriedAutoConnect || Device == null || Device.IsConnected)
                return;

            string? savedPort = Device.LoadSavedComPort();
            if (string.IsNullOrEmpty(savedPort)
                || ComPortComboBox.SelectedItem as string != savedPort)
                return;

            _hasTriedAutoConnect = true;
            _ = DeferredAutoConnectAsync(savedPort);
        }

        private async System.Threading.Tasks.Task DeferredAutoConnectAsync(string port)
        {
            try
            {
                await System.Threading.Tasks.Task.Delay(300);
                if (Device == null || Device.IsConnected || _isSerialBusy)
                    return;

                _isSerialBusy = true;
                UpdateConnectionUi();
                bool ok = await System.Threading.Tasks.Task.Run(() => Device.Connect(port));
                if (ok)
                    Device.StartVolumeMonitoring();
                _isSerialBusy = false;
                UpdateConnectionUi();
                if (!ok && XamlRoot != null)
                    await ShowMessageAsync("接続失敗", $"ポート {port} への接続に失敗しました。");
            }
            catch (Exception ex)
            {
                _isSerialBusy = false;
                UpdateConnectionUi();
                System.Diagnostics.Debug.WriteLine("DeferredAutoConnect: " + ex.Message);
            }
        }

        // ---------- ボタンハンドラ ----------

        private async void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            if (Device == null || _isSerialBusy) return;

            if (Device.IsConnected)
            {
                _isSerialBusy = true;
                UpdateConnectionUi();
                try
                {
                    // 切断の Close 待ちは UI を固めるので BG へ
                    await System.Threading.Tasks.Task.Run(() => Device.Disconnect());
                }
                finally
                {
                    _isSerialBusy = false;
                    UpdateConnectionUi();
                }
                return;
            }

            await TryConnectSelectedPortAsync();
        }

        private async System.Threading.Tasks.Task TryConnectSelectedPortAsync()
        {
            if (Device == null || _isSerialBusy) return;

            string? port = ComPortComboBox.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(port))
            {
                await ShowMessageAsync("接続エラー", "COMポートを選択してください。");
                return;
            }

            _isSerialBusy = true;
            UpdateConnectionUi();
            try
            {
                bool ok = await System.Threading.Tasks.Task.Run(() => Device.Connect(port));
                if (ok)
                    Device.StartVolumeMonitoring();
                UpdateConnectionUi();
                if (!ok)
                    await ShowMessageAsync("接続失敗", $"ポート {port} への接続に失敗しました。");
            }
            finally
            {
                _isSerialBusy = false;
                UpdateConnectionUi();
            }
        }

        private async void AddPatternButton_Click(object sender, RoutedEventArgs e)
        {
            if (Device == null) return;

            var added = Device.AddPattern();
            if (added == null)
            {
                if (Device.Patterns.Count >= 30)
                    await ShowMessageAsync("上限", "パターンの登録上限に達しています。");
                else
                    await ShowMessageAsync(
                        "追加不可",
                        "全ての組み合わせが使用されています。既存のパターンを削除してから追加してください。");
                return;
            }

            // PatternsChanged で再描画されるが、念のため
            RenderAllPatterns();
        }

        private async void SyncAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (Device == null) return;

            if (!Device.IsConnected)
            {
                await ShowMessageAsync("未接続", "マイコンに接続してから同期してください。");
                return;
            }

            bool ok = Device.SyncAllToPicoAndFlash();
            if (!ok)
            {
                await ShowMessageAsync(
                    "重複エラー",
                    "設定の重複があります。同じボタンの組み合わせが複数あります。\n修正してから再度保存・同期してください。");
            }
        }

        private async System.Threading.Tasks.Task ShowMessageAsync(string title, string content)
        {
            if (XamlRoot == null) return;
            var dialog = new ContentDialog
            {
                Title = title,
                Content = content,
                CloseButtonText = "OK",
                XamlRoot = XamlRoot
            };
            await dialog.ShowAsync();
        }

        // ---------- パターン描画 ----------

        private void RenderAllPatterns()
        {
            if (Device == null) return;

            _isRenderingPatterns = true;
            try
            {
                PatternsConfigPanel.Children.Clear();

                var patterns = Device.Patterns;
                for (int i = 0; i < patterns.Count; i++)
                {
                    var pat = patterns[i];
                    if (!IsPatternVisible(pat))
                        continue;
                    RenderPatternCard(pat, i);
                }
            }
            finally
            {
                _isRenderingPatterns = false;
                // 内容増減後に PageScrollHost のスクロール要否を再計算
                ScrollHost.InvalidateScrollability();
            }
        }

        private bool IsPatternVisible(PatternMacroConfig pat)
        {
            if (Device == null) return false;
            int activeCount = Device.ActiveButtonCount;

            if (pat.TriggerType == 3)
                return pat.TriggerParam1 == 1;

            if (pat.TriggerType == 0 && pat.TriggerParam1 > activeCount)
                return false;
            if (pat.TriggerParam1 > activeCount)
                return false;
            if (pat.TriggerType == 1 && pat.TriggerParam2 > activeCount)
                return false;
            return true;
        }

        private void OnPatternBlocksReordered()
        {
            if (Device == null || _isRenderingPatterns) return;

            object?[] tags = DragReorderHelper.GetOrderedTags(PatternsConfigPanel);
            var newVisible = new System.Collections.Generic.List<PatternMacroConfig>();
            foreach (object? tag in tags)
            {
                if (tag is PatternMacroConfig p)
                    newVisible.Add(p);
            }

            if (newVisible.Count == 0) return;

            // 非表示パターンの相対位置を保ちつつ、表示中ブロックの新順序を反映
            var original = Device.Patterns;
            int vi = 0;
            var merged = new System.Collections.Generic.List<PatternMacroConfig>(original.Count);
            foreach (var p in original)
            {
                if (IsPatternVisible(p))
                {
                    if (vi < newVisible.Count)
                        merged.Add(newVisible[vi++]);
                }
                else
                {
                    merged.Add(p);
                }
            }

            while (vi < newVisible.Count)
                merged.Add(newVisible[vi++]);

            bool changed = false;
            if (merged.Count == original.Count)
            {
                for (int i = 0; i < merged.Count; i++)
                {
                    if (!ReferenceEquals(merged[i], original[i]))
                    {
                        changed = true;
                        break;
                    }
                }
            }
            else
            {
                changed = true;
            }

            // UI は既に Children を入れ替え済みなので再描画しない（チラつき防止）
            if (changed)
                Device.ReorderPatterns(merged, scheduleSync: false, raiseChanged: false);
        }

        private void RenderPatternCard(PatternMacroConfig pattern, int index)
        {
            if (Device == null) return;

            int activeCount = Device.ActiveButtonCount;
            bool isBasePattern = pattern.TriggerType == 0 && pattern.TriggerParam1 <= activeCount;
            bool isBaseVolume = pattern.TriggerType == 3 && pattern.TriggerParam1 == 1;
            bool isUndeletable = isBasePattern || isBaseVolume;

            var card = new Border
            {
                Tag = pattern,
                Padding = new Thickness(0),
                CornerRadius = new CornerRadius(8),
                BorderThickness = new Thickness(1),
                Background = (Brush)Application.Current.Resources["CardBackgroundFillColorDefaultBrush"],
                BorderBrush = (Brush)Application.Current.Resources["CardStrokeColorDefaultBrush"],
                RenderTransform = new TranslateTransform()
            };

            // 本体（左） + ☰（右・ブロック高さ中央）
            var cardRoot = new Grid();
            cardRoot.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            cardRoot.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(44) });

            var container = new StackPanel { Spacing = 12, Margin = new Thickness(16, 16, 8, 16) };

            // --- ヘッダー（名前 + 削除）※☰は右列に分離 ---
            var headerPanel = new Grid();
            headerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var headTitle = new TextBox
            {
                Text = string.IsNullOrWhiteSpace(pattern.Name)
                    ? Device.GenerateAutoName(pattern)
                    : pattern.Name,
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
                FontSize = 16,
                Width = 220,
                HorizontalAlignment = HorizontalAlignment.Left
            };
            headTitle.TextChanged += (s, e) =>
            {
                pattern.Name = headTitle.Text;
                ScheduleAutoSync(pattern);
            };
            Grid.SetColumn(headTitle, 0);

            Action triggerChanged = () =>
            {
                if (Device.IsDefaultName(pattern.Name))
                {
                    pattern.Name = Device.GenerateAutoName(pattern);
                    headTitle.Text = pattern.Name;
                }
                ScheduleAutoSync(pattern);
            };

            var deleteBtn = new Button
            {
                Content = "✕",
                Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 244, 67, 54)),
                MinWidth = 36,
                VerticalAlignment = VerticalAlignment.Center,
                Visibility = isUndeletable ? Visibility.Collapsed : Visibility.Visible
            };
            ToolTipService.SetToolTip(deleteBtn, "削除");
            Grid.SetColumn(deleteBtn, 1);
            deleteBtn.Click += async (s, e) =>
            {
                if (XamlRoot == null) return;
                var confirm = new ContentDialog
                {
                    Title = "削除確認",
                    Content = "このパターンを削除しますか？",
                    PrimaryButtonText = "削除",
                    CloseButtonText = "キャンセル",
                    DefaultButton = ContentDialogButton.Close,
                    XamlRoot = XamlRoot
                };
                if (await confirm.ShowAsync() != ContentDialogResult.Primary)
                    return;
                Device.DeletePattern(pattern);
            };

            headerPanel.Children.Add(headTitle);
            headerPanel.Children.Add(deleteBtn);
            container.Children.Add(headerPanel);

            // --- トリガー行 ---
            var triggerPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 12
            };

            var tTypeCombo = new ComboBox { Width = 150, MinWidth = 120 };
            if (isBasePattern)
            {
                tTypeCombo.Items.Add(MakeTagItem("単押し", 0));
                tTypeCombo.SelectedIndex = 0;
                tTypeCombo.IsEnabled = false;
            }
            else if (isBaseVolume)
            {
                tTypeCombo.Items.Add(MakeTagItem("ボリューム", 3));
                tTypeCombo.SelectedIndex = 0;
                tTypeCombo.IsEnabled = false;
            }
            else
            {
                tTypeCombo.Items.Add(MakeTagItem("同時押し", 1));
                tTypeCombo.Items.Add(MakeTagItem("複数回押し", 2));
                if (pattern.TriggerType == 0 || pattern.TriggerType == 3)
                    pattern.TriggerType = 1;
                if (pattern.TriggerType == 1) tTypeCombo.SelectedIndex = 0;
                else if (pattern.TriggerType == 2) tTypeCombo.SelectedIndex = 1;
                else tTypeCombo.SelectedIndex = 0;
            }

            var tParam1Combo = new ComboBox { Width = 110, MinWidth = 100 };
            if (isBasePattern)
            {
                tParam1Combo.Items.Add(MakeTagItem($"ボタン{pattern.TriggerParam1}", pattern.TriggerParam1));
                tParam1Combo.SelectedIndex = 0;
                tParam1Combo.IsEnabled = false;
            }
            else if (isBaseVolume)
            {
                tParam1Combo.Items.Add(MakeTagItem("ボリューム", 1));
                tParam1Combo.SelectedIndex = 0;
                tParam1Combo.IsEnabled = false;
            }
            else
            {
                for (int i = 1; i <= activeCount; i++)
                    tParam1Combo.Items.Add(MakeTagItem($"ボタン{i}", i));
                SetComboByTag(tParam1Combo, pattern.TriggerParam1);
                if (tParam1Combo.SelectedIndex < 0) tParam1Combo.SelectedIndex = 0;
            }

            var param2Panel = new StackPanel();
            var tParam2Combo = new ComboBox { Width = 110, MinWidth = 100 };
            param2Panel.Children.Add(tParam2Combo);

            var repeatTxt = new TextBox
            {
                Width = 120,
                Text = pattern.RepeatInterval.ToString(),
                PlaceholderText = "連続間隔(ms)"
            };
            repeatTxt.TextChanged += (s, e) =>
            {
                if (int.TryParse(repeatTxt.Text, out int v))
                {
                    pattern.RepeatInterval = v;
                    ScheduleAutoSync(pattern);
                }
            };

            Action updateParam2Ui = () =>
            {
                int pType = pattern.TriggerType;
                if (pType == 0)
                {
                    param2Panel.Visibility = Visibility.Collapsed;
                }
                else if (pType == 1)
                {
                    param2Panel.Visibility = Visibility.Visible;
                    int oldVal = pattern.TriggerParam2;
                    tParam2Combo.Items.Clear();
                    for (int i = 1; i <= activeCount; i++)
                    {
                        if (i != pattern.TriggerParam1)
                            tParam2Combo.Items.Add(MakeTagItem($"ボタン{i}", i));
                    }
                    SetComboByTag(tParam2Combo, oldVal);
                    if (tParam2Combo.SelectedIndex < 0 && tParam2Combo.Items.Count > 0)
                        tParam2Combo.SelectedIndex = 0;
                }
                else if (pType == 2)
                {
                    param2Panel.Visibility = Visibility.Visible;
                    int oldVal = pattern.TriggerParam2;
                    tParam2Combo.Items.Clear();
                    tParam2Combo.Items.Add(MakeTagItem("2回", 2));
                    tParam2Combo.Items.Add(MakeTagItem("3回", 3));
                    SetComboByTag(tParam2Combo, oldVal);
                    if (tParam2Combo.SelectedIndex < 0) tParam2Combo.SelectedIndex = 0;
                }

                if (GetComboTagInt(tParam2Combo) is int p2)
                    pattern.TriggerParam2 = p2;
            };

            updateParam2Ui();

            bool isUpdating = false;

            tTypeCombo.SelectionChanged += async (s, e) =>
            {
                if (isUpdating || tTypeCombo.SelectedItem == null) return;
                int oldType = pattern.TriggerType;
                int oldP2 = pattern.TriggerParam2;
                int? newType = GetComboTagInt(tTypeCombo);
                if (newType == null || oldType == newType.Value) return;

                pattern.TriggerType = newType.Value;
                if (newType == 1 && pattern.TriggerParam1 == pattern.TriggerParam2)
                    pattern.TriggerParam2 = pattern.TriggerParam1 == 1 ? 2 : 1;
                else if (newType == 2 && pattern.TriggerParam2 < 2)
                    pattern.TriggerParam2 = 2;

                if (!isBasePattern && Device.CheckDuplicate(pattern))
                {
                    await ShowMessageAsync(
                        "重複エラー",
                        "すでに同じ組み合わせのパターンが存在します。別の組み合わせを選択してください。");
                    pattern.TriggerType = oldType;
                    pattern.TriggerParam2 = oldP2;
                    isUpdating = true;
                    SetComboByTag(tTypeCombo, oldType);
                    isUpdating = false;
                    return;
                }

                isUpdating = true;
                updateParam2Ui();
                triggerChanged();
                isUpdating = false;
            };

            tParam1Combo.SelectionChanged += async (s, e) =>
            {
                if (isUpdating || tParam1Combo.SelectedItem == null) return;
                int oldP1 = pattern.TriggerParam1;
                int? newP1 = GetComboTagInt(tParam1Combo);
                if (newP1 == null || oldP1 == newP1.Value) return;

                pattern.TriggerParam1 = newP1.Value;
                int oldP2 = pattern.TriggerParam2;
                if (pattern.TriggerType == 1 && newP1 == pattern.TriggerParam2)
                    pattern.TriggerParam2 = newP1 == 1 ? 2 : 1;

                if (!isBasePattern && Device.CheckDuplicate(pattern))
                {
                    await ShowMessageAsync(
                        "重複エラー",
                        "すでに同じ組み合わせのパターンが存在します。別のボタンを選択してください。");
                    pattern.TriggerParam1 = oldP1;
                    pattern.TriggerParam2 = oldP2;
                    isUpdating = true;
                    SetComboByTag(tParam1Combo, oldP1);
                    isUpdating = false;
                    return;
                }

                isUpdating = true;
                updateParam2Ui();
                triggerChanged();
                isUpdating = false;
            };

            tParam2Combo.SelectionChanged += async (s, e) =>
            {
                if (isUpdating || tParam2Combo.SelectedItem == null) return;
                int oldP2 = pattern.TriggerParam2;
                int? newP2 = GetComboTagInt(tParam2Combo);
                if (newP2 == null || oldP2 == newP2.Value) return;

                pattern.TriggerParam2 = newP2.Value;

                if (!isBasePattern && Device.CheckDuplicate(pattern))
                {
                    await ShowMessageAsync(
                        "重複エラー",
                        "すでに同じ組み合わせのパターンが存在します。別の値を選択してください。");
                    pattern.TriggerParam2 = oldP2;
                    isUpdating = true;
                    SetComboByTag(tParam2Combo, oldP2);
                    isUpdating = false;
                    return;
                }

                isUpdating = true;
                triggerChanged();
                isUpdating = false;
            };

            triggerPanel.Children.Add(tTypeCombo);
            triggerPanel.Children.Add(tParam1Combo);
            if (!isBaseVolume)
            {
                triggerPanel.Children.Add(param2Panel);
                triggerPanel.Children.Add(repeatTxt);
            }
            container.Children.Add(triggerPanel);

            if (isBaseVolume)
            {
                container.Children.Add(new TextBlock
                {
                    Text = "音量はロータリーエンコーダ（24クリック）で調節します。1クリックあたり±2%（接続時はアプリ経由、未接続時はHID）。",
                    Opacity = 0.7,
                    TextWrapping = TextWrapping.WrapWholeWords
                });
            }

            // --- マウス一括登録ボタン（フッター側で配置） ---
            var mouseCapMainBtn = new Button
            {
                HorizontalAlignment = HorizontalAlignment.Left
            };

            Action updatePatternMouseBtnVisibility = () =>
            {
                bool hasMouse = pattern.Steps.Any(st => st.Type == "MOUSE");
                mouseCapMainBtn.Visibility = hasMouse ? Visibility.Visible : Visibility.Collapsed;

                bool capturingThis = _mouseCapture != null
                    && _mouseCapture.IsCapturing
                    && ReferenceEquals(_mouseCapture.CapturingPattern, pattern);

                if (capturingThis)
                {
                    mouseCapMainBtn.Content = _mouseCapture!.CaptureCount > 0
                        ? $"登録中 ({_mouseCapture.CaptureCount})..."
                        : "一括登録 停止 (Esc)";
                    mouseCapMainBtn.Foreground = new SolidColorBrush(
                        Windows.UI.Color.FromArgb(255, 244, 67, 54));
                }
                else
                {
                    mouseCapMainBtn.Content = "マウス一括登録";
                    mouseCapMainBtn.ClearValue(Control.ForegroundProperty);
                }
            };

            // --- ステップ一覧（Grid + ☰ DnD） ---
            const double stepItemHeight = 48.0;
            var stepsContainerGrid = new Grid();
            for (int i = 0; i < pattern.Steps.Count; i++)
                stepsContainerGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(stepItemHeight) });

            for (int i = 0; i < pattern.Steps.Count; i++)
            {
                int stepIndex = i;
                var step = pattern.Steps[stepIndex];
                var stepRow = BuildStepRow(pattern, step, stepIndex, updatePatternMouseBtnVisibility, stepsContainerGrid, stepItemHeight);
                Grid.SetRow(stepRow, stepIndex);
                stepsContainerGrid.Children.Add(stepRow);
            }

            container.Children.Add(stepsContainerGrid);

            // --- フッター ---
            var footerPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 12
            };

            var addStepBtn = new Button { Content = "＋ ステップ追加" };
            addStepBtn.Click += async (s, e) =>
            {
                if (pattern.Steps.Count >= 10)
                {
                    await ShowMessageAsync("上限", "ステップ上限(10)です。");
                    return;
                }
                pattern.Steps.Add(new MacroStepConfig { Type = "KEY", Data = "" });
                Device.SavePatterns();
                ScheduleAutoSync(pattern);
                RenderAllPatterns();
            };
            if (pattern.Steps.Count < 10)
                footerPanel.Children.Add(addStepBtn);

            mouseCapMainBtn.Click += (s, e) =>
            {
                if (_mouseCapture != null && _mouseCapture.IsCapturing)
                {
                    _mouseCapture.Stop();
                    return;
                }

                int startIndex = pattern.Steps.Count;
                for (int i = 0; i < pattern.Steps.Count; i++)
                {
                    if (pattern.Steps[i].Type == "MOUSE"
                        && string.IsNullOrEmpty(pattern.Steps[i].Data))
                    {
                        startIndex = i;
                        break;
                    }
                }

                EnsureMouseCapture();
                _mouseCapture!.Start(pattern, startIndex);
            };
            footerPanel.Children.Add(mouseCapMainBtn);
            updatePatternMouseBtnVisibility();

            container.Children.Add(footerPanel);

            Grid.SetColumn(container, 0);
            cardRoot.Children.Add(container);

            var blockDragHandle = new TextBlock
            {
                Text = "☰",
                FontSize = 22,
                FontWeight = Microsoft.UI.Text.FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Opacity = 0.55
            };
            ToolTipService.SetToolTip(blockDragHandle, "ドラッグでブロックを並び替え");
            blockDragHandle.PointerEntered += (s, e) => { if (s is TextBlock tb) tb.Opacity = 1.0; };
            blockDragHandle.PointerExited += (s, e) => { if (s is TextBlock tb) tb.Opacity = 0.55; };
            Grid.SetColumn(blockDragHandle, 1);
            cardRoot.Children.Add(blockDragHandle);

            card.Child = cardRoot;
            PatternsConfigPanel.Children.Add(card);

            DragReorderHelper.AttachToStackPanel(
                blockDragHandle,
                card,
                PatternsConfigPanel,
                OnPatternBlocksReordered);
        }

        private FrameworkElement BuildStepRow(
            PatternMacroConfig pattern,
            MacroStepConfig step,
            int stepIndex,
            Action updatePatternMouseBtnVisibility,
            Grid stepsContainerGrid,
            double stepItemHeight)
        {
            var stepGrid = new Grid
            {
                Tag = step,
                Margin = new Thickness(0, 2, 0, 2),
                RenderTransform = new TranslateTransform()
            };
            stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(28) }); // ☰
            stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) }); // 番号
            stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(130) });
            stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var dragHandle = new TextBlock
            {
                Text = "☰",
                FontSize = 16,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Opacity = 0.55
            };
            ToolTipService.SetToolTip(dragHandle, "ドラッグで並び替え");
            Grid.SetColumn(dragHandle, 0);

            var numBlock = new TextBlock
            {
                Text = $"{stepIndex + 1}.",
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            };
            Grid.SetColumn(numBlock, 1);

            var typeCombo = new ComboBox
            {
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center,
                MinWidth = 120
            };
            typeCombo.Items.Add(MakeTagItem("キーボード", "KEY"));
            typeCombo.Items.Add(MakeTagItem("マウス座標", "MOUSE"));
            typeCombo.Items.Add(MakeTagItem("アプリ起動", "CMD"));
            typeCombo.Items.Add(MakeTagItem("待機", "WAIT"));
            SetComboByTag(typeCombo, step.Type);
            Grid.SetColumn(typeCombo, 2);

            var inputTxt = new TextBox
            {
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Text = step.Data,
                MinWidth = 120
            };
            Grid.SetColumn(inputTxt, 3);

            var browseBtn = new Button
            {
                Content = "参照",
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(browseBtn, 4);

            var stepDelBtn = new Button
            {
                Content = "✕",
                Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 244, 67, 54)),
                MinWidth = 36,
                VerticalAlignment = VerticalAlignment.Center,
                Visibility = stepIndex == 0 ? Visibility.Collapsed : Visibility.Visible
            };
            Grid.SetColumn(stepDelBtn, 5);

            Action updateStepUi = () =>
            {
                if (step.Type == "KEY")
                    inputTxt.PlaceholderText = "クリックして入力";
                else if (step.Type == "MOUSE")
                    inputTxt.PlaceholderText = "座標 (例: 16000,8000)";
                else if (step.Type == "CMD")
                    inputTxt.PlaceholderText = "コマンド/EXE";
                else if (step.Type == "WAIT")
                    inputTxt.PlaceholderText = "待機ミリ秒 (例: 500)";

                inputTxt.IsReadOnly = step.Type == "KEY";
                browseBtn.Visibility = step.Type == "CMD" ? Visibility.Visible : Visibility.Collapsed;
            };

            typeCombo.SelectionChanged += (s, e) =>
            {
                string? tag = GetComboTagString(typeCombo);
                if (tag == null) return;
                step.Type = tag;
                step.Data = "";
                inputTxt.Text = "";
                updateStepUi();
                updatePatternMouseBtnVisibility();
                ScheduleAutoSync(pattern);
            };

            inputTxt.TextChanged += (s, e) =>
            {
                step.Data = inputTxt.Text;
                ScheduleAutoSync(pattern);
            };

            inputTxt.KeyDown += (s, e) =>
            {
                if (step.Type != "KEY") return;
                e.Handled = true;

                VirtualKey actualKey = e.Key;
                if (actualKey == VirtualKey.Control
                    || actualKey == VirtualKey.LeftControl
                    || actualKey == VirtualKey.RightControl
                    || actualKey == VirtualKey.Shift
                    || actualKey == VirtualKey.LeftShift
                    || actualKey == VirtualKey.RightShift
                    || actualKey == VirtualKey.Menu
                    || actualKey == VirtualKey.LeftMenu
                    || actualKey == VirtualKey.RightMenu
                    || actualKey == VirtualKey.LeftWindows
                    || actualKey == VirtualKey.RightWindows)
                {
                    return;
                }

                string modifiers = "";
                if (IsKeyDown(VirtualKey.Control) || IsKeyDown(VirtualKey.LeftControl) || IsKeyDown(VirtualKey.RightControl))
                    modifiers += "Ctrl+";
                if (IsKeyDown(VirtualKey.Shift) || IsKeyDown(VirtualKey.LeftShift) || IsKeyDown(VirtualKey.RightShift))
                    modifiers += "Shift+";
                if (IsKeyDown(VirtualKey.Menu) || IsKeyDown(VirtualKey.LeftMenu) || IsKeyDown(VirtualKey.RightMenu))
                    modifiers += "Alt+";

                string keyStr = actualKey.ToString();

                if (actualKey >= VirtualKey.A && actualKey <= VirtualKey.Z)
                {
                    bool shift = IsKeyDown(VirtualKey.Shift)
                        || IsKeyDown(VirtualKey.LeftShift)
                        || IsKeyDown(VirtualKey.RightShift);
                    if (!shift)
                        keyStr = keyStr.ToLowerInvariant();
                }

                if (actualKey >= VirtualKey.Number0 && actualKey <= VirtualKey.Number9)
                    keyStr = ((int)(actualKey - VirtualKey.Number0)).ToString();

                if (actualKey == VirtualKey.Enter) keyStr = "Enter";
                if (actualKey == VirtualKey.Escape) keyStr = "Esc";
                if (actualKey == VirtualKey.Space) keyStr = "Space";

                inputTxt.Text = modifiers + keyStr;
            };

            browseBtn.Click += async (s, e) =>
            {
                if (XamlRoot == null) return;
                string? path = await AppSelectorDialog.ShowAsync(XamlRoot);
                if (!string.IsNullOrEmpty(path))
                    inputTxt.Text = path;
            };

            stepDelBtn.Click += (s, e) =>
            {
                int idx = pattern.Steps.IndexOf(step);
                if (idx <= 0) return;
                pattern.Steps.RemoveAt(idx);
                Device?.SavePatterns();
                ScheduleAutoSync(pattern);
                RenderAllPatterns();
            };

            updateStepUi();

            stepGrid.Children.Add(dragHandle);
            stepGrid.Children.Add(numBlock);
            stepGrid.Children.Add(typeCombo);
            stepGrid.Children.Add(inputTxt);
            stepGrid.Children.Add(browseBtn);
            stepGrid.Children.Add(stepDelBtn);

            DragReorderHelper.Attach(
                dragHandle,
                stepGrid,
                stepsContainerGrid,
                stepItemHeight,
                () =>
                {
                    object?[] tags = DragReorderHelper.GetOrderedTags(stepsContainerGrid);
                    var newOrder = new List<MacroStepConfig>();
                    foreach (object? tag in tags)
                    {
                        if (tag is MacroStepConfig cfg)
                            newOrder.Add(cfg);
                    }

                    if (newOrder.Count != pattern.Steps.Count)
                        return;

                    bool changed = false;
                    for (int i = 0; i < newOrder.Count; i++)
                    {
                        if (!ReferenceEquals(newOrder[i], pattern.Steps[i]))
                        {
                            changed = true;
                            break;
                        }
                    }

                    if (!changed)
                        return;

                    pattern.Steps.Clear();
                    pattern.Steps.AddRange(newOrder);

                    // 番号ラベル更新
                    foreach (UIElement child in stepsContainerGrid.Children)
                    {
                        if (child is Grid g && g.Tag is MacroStepConfig cfg)
                        {
                            int row = pattern.Steps.IndexOf(cfg);
                            Grid.SetRow(g, row);
                            var num = g.Children.OfType<TextBlock>()
                                .FirstOrDefault(t => Grid.GetColumn(t) == 1);
                            if (num != null && num.Text.EndsWith("."))
                                num.Text = $"{row + 1}.";

                            // 先頭ステップの削除ボタンは非表示
                            var del = g.Children.OfType<Button>()
                                .FirstOrDefault(b => b.Content as string == "✕");
                            if (del != null)
                                del.Visibility = row == 0 ? Visibility.Collapsed : Visibility.Visible;
                        }
                    }

                    Device?.SavePatterns();
                    ScheduleAutoSync(pattern);
                });

            return stepGrid;
        }

        private void EnsureMouseCapture()
        {
            if (_mouseCapture != null) return;
            _mouseCapture = new MouseCaptureHelper(
                DispatcherQueue,
                onChanged: () =>
                {
                    Device?.SavePatterns();
                    RenderAllPatterns();
                });
        }

        private void ScheduleAutoSync(PatternMacroConfig? changedPattern = null)
        {
            if (_isRenderingPatterns) return;
            Device?.ScheduleAutoSync(changedPattern);
        }

        // ---------- ComboBox ヘルパー ----------

        private static ComboBoxItem MakeTagItem(string content, object tag)
        {
            return new ComboBoxItem { Content = content, Tag = tag };
        }

        private static void SetComboByTag(ComboBox cb, object tag)
        {
            string t = tag.ToString() ?? "";
            for (int i = 0; i < cb.Items.Count; i++)
            {
                if (cb.Items[i] is ComboBoxItem item
                    && (item.Tag?.ToString() ?? "") == t)
                {
                    cb.SelectedIndex = i;
                    return;
                }
            }
        }

        private static int? GetComboTagInt(ComboBox cb)
        {
            if (cb.SelectedItem is ComboBoxItem item && item.Tag is int i)
                return i;
            if (cb.SelectedItem is ComboBoxItem item2
                && int.TryParse(item2.Tag?.ToString(), out int parsed))
                return parsed;
            return null;
        }

        private static string? GetComboTagString(ComboBox cb)
        {
            if (cb.SelectedItem is ComboBoxItem item)
                return item.Tag?.ToString();
            return null;
        }

        private static bool IsKeyDown(VirtualKey key)
        {
            var state = InputKeyboardSource.GetKeyStateForCurrentThread(key);
            return (state & CoreVirtualKeyStates.Down) == CoreVirtualKeyStates.Down;
        }
    }

    /// <summary>
    /// 低レベルマウス／キーボードフックによる座標一括キャプチャ。
    /// Esc で停止。座標は HID 正規化 (0–32767)。
    /// </summary>
    internal sealed class MouseCaptureHelper : IDisposable
    {
        private const int WhMouseLl = 14;
        private const int WhKeyboardLl = 13;
        private const int WmLbuttonDown = 0x0201;
        private const int WmKeyDown = 0x0100;
        private const int VkEscape = 0x1B;

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private readonly Microsoft.UI.Dispatching.DispatcherQueue _dispatcher;
        private readonly Action _onChanged;
        private readonly ClickMarkerOverlay _clickMarkers = new();

        private LowLevelMouseProc? _mouseProc;
        private LowLevelKeyboardProc? _keyboardProc;
        private IntPtr _mouseHookId = IntPtr.Zero;
        private IntPtr _keyboardHookId = IntPtr.Zero;

        private bool _isCapturing;
        private PatternMacroConfig? _capturingPattern;
        private int _capturingStepIndex = -1;
        private int _captureCount;

        public bool IsCapturing => _isCapturing;
        public PatternMacroConfig? CapturingPattern => _capturingPattern;
        public int CaptureCount => _captureCount;

        public MouseCaptureHelper(Microsoft.UI.Dispatching.DispatcherQueue dispatcher, Action onChanged)
        {
            _dispatcher = dispatcher;
            _onChanged = onChanged;
            _mouseProc = MouseHookCallback;
            _keyboardProc = KeyboardHookCallback;
        }

        public void Start(PatternMacroConfig pattern, int startStepIndex)
        {
            if (pattern.Steps.Count >= 10 && startStepIndex >= 10)
                return;

            Stop();

            _capturingPattern = pattern;
            _capturingStepIndex = startStepIndex;
            _isCapturing = true;
            _captureCount = 0;

            using Process curProcess = Process.GetCurrentProcess();
            using ProcessModule? curModule = curProcess.MainModule;
            if (curModule == null) return;

            IntPtr handle = GetModuleHandle(curModule.ModuleName);
            _mouseHookId = SetWindowsHookEx(WhMouseLl, _mouseProc!, handle, 0);
            _keyboardHookId = SetWindowsHookEx(WhKeyboardLl, _keyboardProc!, handle, 0);

            _onChanged();
        }

        public void Stop()
        {
            if (!_isCapturing) return;

            _clickMarkers.Clear();

            if (_mouseHookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_mouseHookId);
                _mouseHookId = IntPtr.Zero;
            }
            if (_keyboardHookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_keyboardHookId);
                _keyboardHookId = IntPtr.Zero;
            }

            _isCapturing = false;
            _capturingPattern = null;
            _capturingStepIndex = -1;

            _dispatcher.TryEnqueue(() => _onChanged());
        }

        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WmLbuttonDown && _isCapturing && _capturingPattern != null)
            {
                var hookStruct = Marshal.PtrToStructure<Msllhookstruct>(lParam);

                int physicalWidth = GetSystemMetrics(0);  // SM_CXSCREEN
                int physicalHeight = GetSystemMetrics(1); // SM_CYSCREEN
                if (physicalWidth <= 0) physicalWidth = 1;
                if (physicalHeight <= 0) physicalHeight = 1;

                int hidX = (int)((hookStruct.pt.x / (double)physicalWidth) * 32767.0);
                int hidY = (int)((hookStruct.pt.y / (double)physicalHeight) * 32767.0);
                hidX = Math.Clamp(hidX, 0, 32767);
                hidY = Math.Clamp(hidY, 0, 32767);
                string data = $"{hidX},{hidY}";

                var pattern = _capturingPattern;
                _dispatcher.TryEnqueue(() =>
                {
                    if (!_isCapturing || pattern == null) return;

                    _captureCount++;
                    _clickMarkers.Show(hookStruct.pt.x, hookStruct.pt.y, _captureCount);

                    if (_captureCount == 1 && _capturingStepIndex < pattern.Steps.Count)
                    {
                        pattern.Steps[_capturingStepIndex].Type = "MOUSE";
                        pattern.Steps[_capturingStepIndex].Data = data;
                    }
                    else if (pattern.Steps.Count < 10)
                    {
                        pattern.Steps.Add(new MacroStepConfig { Type = "MOUSE", Data = data });
                        if (pattern.Steps.Count >= 10)
                        {
                            Stop();
                            return;
                        }
                    }
                    else
                    {
                        Stop();
                        return;
                    }

                    _onChanged();
                });
            }

            return CallNextHookEx(_mouseHookId, nCode, wParam, lParam);
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WmKeyDown)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                if (vkCode == VkEscape)
                    _dispatcher.TryEnqueue(Stop);
            }
            return CallNextHookEx(_keyboardHookId, nCode, wParam, lParam);
        }

        public void Dispose()
        {
            Stop();
            _clickMarkers.Dispose();
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, Delegate lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        [StructLayout(LayoutKind.Sequential)]
        private struct Point
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Msllhookstruct
        {
            public Point pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }
    }
}
