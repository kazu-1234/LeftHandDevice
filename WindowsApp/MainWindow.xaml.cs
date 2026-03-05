using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Media;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Threading.Tasks;
using System.Net.Http;
using MaterialDesignThemes.Wpf;

namespace LeftHandDeviceApp
{
    public class MacroStepConfig
    {
        public string Type { get; set; } = "KEY"; // KEY, MOUSE, CMD, WAIT
        public string Data { get; set; } = "";
    }

    public class PatternMacroConfig
    {
        public string Name { get; set; } = "";
        public int TriggerType { get; set; } = 0; // 0=単押し, 1=同時押し, 2=複数回押し
        public int TriggerParam1 { get; set; } = 1; // 対象ボタン (1〜5)
        public int TriggerParam2 { get; set; } = 2; // 同時押しのボタン2(1〜5)、または複数回押しの回数(2〜3)
        public int RepeatInterval { get; set; } = 200; // 連続間隔 (ms)
        public List<MacroStepConfig> Steps { get; set; } = new List<MacroStepConfig>();
    }

    public partial class MainWindow : Window
    {
        private SerialPort _serialPort;
        private bool _isConnected = false;
        
        private List<PatternMacroConfig> _patterns = new List<PatternMacroConfig>();
        private int _activeButtonCount = 5; // 初期値

        private readonly string _comPortFilePath = Path.Combine(Path.GetDirectoryName(Environment.ProcessPath ?? AppDomain.CurrentDomain.BaseDirectory), "saved_com_port.txt");
        private readonly string _settingsFilePath = Path.Combine(Path.GetDirectoryName(Environment.ProcessPath ?? AppDomain.CurrentDomain.BaseDirectory), "app_settings.json");
        private readonly string _patternsFilePath = Path.Combine(Path.GetDirectoryName(Environment.ProcessPath ?? AppDomain.CurrentDomain.BaseDirectory), "app_patterns.json");

        // --- Hooks ---
        private const int WH_MOUSE_LL = 14;
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_KEYDOWN = 0x0100;
        private const int VK_ESCAPE = 0x1B;

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private LowLevelMouseProc _mouseProc;
        private LowLevelKeyboardProc _keyboardProc;

        private IntPtr _mouseHookID = IntPtr.Zero;
        private IntPtr _keyboardHookID = IntPtr.Zero;

        private bool _isCapturing = false;
        private PatternMacroConfig _capturingPattern = null;
        private int _capturingStepIndex = -1;
        private int _captureCount = 0;
        private List<Window> _overlayMarkers = new List<Window>();

        private PatternMacroConfig _lastChangedPattern = null;
        private bool _isRenderingPatterns = false;

        // 連続動作中の状態管理
        private HashSet<int> _continuousActiveButtons = new HashSet<int>();
        private bool _warningSound = true; // デフォルトで警告音ON
        private System.Windows.Threading.DispatcherTimer _warningHideTimer;

        // ステップD&D用
        private int _dragStepSourceIndex = -1;
        private PatternMacroConfig _dragStepPattern = null;
        private DragAdorner _stepDragAdorner = null;
        private HashSet<int> _unsyncedChangedButtons = new HashSet<int>();

        // 自動送信用タイマー（UIの変更を500ms待ってからまとめて送信）
        private System.Windows.Threading.DispatcherTimer _autoSyncTimer;

        private static readonly string[] CircledNumbers = { "❶", "❷", "❸", "❹", "❺", "❻", "❼", "❽", "❾", "❿" };

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, Delegate lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int x; public int y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT { public POINT pt; public uint mouseData; public uint flags; public uint time; public IntPtr dwExtraInfo; }

        public MainWindow()
        {
            InitializeComponent();

            _mouseProc = MouseHookCallback;
            _keyboardProc = KeyboardHookCallback;

            LoadSettings();
            LoadPatterns();

            bool hasValidSavedPort = LoadComPorts();
            RenderAllPatterns();

            // 警告音設定の読み込み
            LoadWarningSoundSetting();

            // 連続動作中の警告表示を一定時間後に自動非表示にするタイマー
            _warningHideTimer = new System.Windows.Threading.DispatcherTimer();
            _warningHideTimer.Interval = TimeSpan.FromSeconds(3);
            _warningHideTimer.Tick += (s, e) =>
            {
                _warningHideTimer.Stop();
                // フェードアウトアニメーション
                var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(300));
                fadeOut.Completed += (s2, e2) =>
                {
                    ContinuousWarningOverlay.Visibility = Visibility.Collapsed;
                    ContinuousWarningOverlay.Opacity = 1;
                };
                ContinuousWarningOverlay.BeginAnimation(OpacityProperty, fadeOut);
            };

            // ウィンドウが非アクティブになった時の警告処理
            this.Deactivated += (s, e) =>
            {
                if (_continuousActiveButtons.Count > 0)
                {
                    ShowContinuousWarning();
                }
            };

            // 自動送信タイマーの初期化（200ms後にまとめて送信）
            _autoSyncTimer = new System.Windows.Threading.DispatcherTimer();
            _autoSyncTimer.Interval = TimeSpan.FromMilliseconds(200);
            _autoSyncTimer.Tick += (s, e) =>
            {
                _autoSyncTimer.Stop();
                SavePatterns();
                SyncAllToPico();

                // データ送信完了後、Arduinoが全コマンドを処理し終えてからLED点滅コマンドを送信
                if (_unsyncedChangedButtons.Count > 0 && _isConnected && _serialPort != null && _serialPort.IsOpen)
                {
                    System.Threading.Thread.Sleep(300);
                    try
                    {
                        foreach (int bIndex in _unsyncedChangedButtons)
                        {
                            if (bIndex >= 0 && bIndex < 5)
                            {
                                _serialPort.WriteLine($"FLASH_BUTTONS:{bIndex}:-1");
                                System.Threading.Thread.Sleep(350); // 連続で送るとマイコン側が処理しきれない場合があるためウェイト
                            }
                        }
                    }
                    catch { }
                }
                _unsyncedChangedButtons.Clear();
            };

            if (ComPortComboBox.Items.Count > 0 && hasValidSavedPort)
            {
                ConnectButton_Click(this, new RoutedEventArgs());
            }

            // 起動時にバックグラウンドでアップデートを確認
            _ = CheckUpdateAtStartupAsync();
        }

        protected override void OnClosed(EventArgs e)
        {
            StopCapture();
            if (_serialPort != null && _serialPort.IsOpen) _serialPort.Close();
            base.OnClosed(e);
        }

        private bool LoadComPorts()
        {
            bool loadedSavedPort = false;
            ComPortComboBox.Items.Clear();
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports) ComPortComboBox.Items.Add(port);
            
            if (ComPortComboBox.Items.Count > 0)
            {
                string savedPort = null;
                if (File.Exists(_comPortFilePath))
                {
                    try { savedPort = File.ReadAllText(_comPortFilePath).Trim(); } catch { }
                }

                if (!string.IsNullOrEmpty(savedPort) && ComPortComboBox.Items.Contains(savedPort))
                {
                    ComPortComboBox.SelectedItem = savedPort;
                    loadedSavedPort = true;
                }
                else
                {
                    ComPortComboBox.SelectedIndex = 0;
                }
            }
            return loadedSavedPort;
        }

        private void LoadSettings()
        {
            if (File.Exists(_settingsFilePath))
            {
                try
                {
                    var json = JObject.Parse(File.ReadAllText(_settingsFilePath));
                    if (json["ActiveButtonCount"] != null)
                        _activeButtonCount = json["ActiveButtonCount"].Value<int>();
                }
                catch { }
            }
            // 有効ボタン数が1~5の範囲外の場合は補正
            _activeButtonCount = Math.Max(1, Math.Min(5, _activeButtonCount));
            ActiveButtonsText.Text = $"有効ボタン数: {_activeButtonCount}";
        }

        private void LoadPatterns()
        {
            if (File.Exists(_patternsFilePath))
            {
                try
                {
                    string json = File.ReadAllText(_patternsFilePath);
                    var loaded = JsonConvert.DeserializeObject<List<PatternMacroConfig>>(json);
                    if (loaded != null) _patterns = loaded;
                }
                catch { }
            }
            // 初回起動時など空の場合、不足分を補う
            if (!File.Exists(_patternsFilePath) && _patterns.Count < 5)
            {
                var existingBtnIds = _patterns.Where(p => p.TriggerType == 0).Select(p => p.TriggerParam1).ToList();
                for (int i = 1; i <= 5; i++)
                {
                    if (!existingBtnIds.Contains(i))
                    {
                        var p = new PatternMacroConfig { TriggerType = 0, TriggerParam1 = i, Name = $"ボタン{i}" };
                        p.Steps.Add(new MacroStepConfig { Type = "KEY", Data = ((char)('a' + (i-1))).ToString() });
                        _patterns.Add(p);
                    }
                }
                // 初回のみボタン番号順に並べ替え
                _patterns = _patterns.OrderBy(p => p.TriggerParam1).ToList();
                SavePatterns();
            }
            // ※設定画面で並び替えた順序を維持するため、ここでは再ソートしない
        }

        public void ReloadPatternsFromSettings()
        {
            // 設定画面から呼ばれた場合、有効ボタン数も再読み込みする
            LoadSettings();
            LoadPatterns();
            RenderAllPatterns();
        }

        private void SavePatterns()
        {
            try
            {
                string json = JsonConvert.SerializeObject(_patterns, Formatting.Indented);
                File.WriteAllText(_patternsFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"パターンの保存に失敗しました: {ex.Message}");
            }
        }

        // --- UI Rendering ---

        private void RenderAllPatterns()
        {
            _isRenderingPatterns = true;

            // 欠損しているベースパターン(単押し)を復元（有効ボタン数の範囲内のみ）
            for (int i = 1; i <= _activeButtonCount; i++)
            {
                if (!_patterns.Any(p => p.TriggerType == 0 && p.TriggerParam1 == i))
                {
                    // 誤って上書きされた可能性があるため、名前に「ボタンi」が含まれるものを優先して単押しに戻す
                    var mutated = _patterns.FirstOrDefault(p => p.Name == $"ボタン{i}");
                    if (mutated != null)
                    {
                        mutated.TriggerType = 0;
                        mutated.TriggerParam1 = i;
                    }
                    else
                    {
                        // なければ新規作成して前方に挿入
                        var newBase = new PatternMacroConfig { TriggerType = 0, TriggerParam1 = i, Name = $"ボタン{i}" };
                        newBase.Steps.Add(new MacroStepConfig { Type = "KEY", Data = "" });
                        _patterns.Insert(Math.Min(i - 1, _patterns.Count), newBase);
                    }
                }
            }

            PatternsConfigPanel.Children.Clear();
            for (int i = 0; i < _patterns.Count; i++)
            {
                var pat = _patterns[i];
                // 有効ボタン数を超えるベースパターン（単押し）は非表示にする
                if (pat.TriggerType == 0 && pat.TriggerParam1 > _activeButtonCount)
                    continue;
                // 有効ボタン数を超えるボタンを使う追加パターンも非表示にする
                if (pat.TriggerParam1 > _activeButtonCount)
                    continue;
                if (pat.TriggerType == 1 && pat.TriggerParam2 > _activeButtonCount)
                    continue;
                RenderPatternCard(i);
            }
            _isRenderingPatterns = false;
        }

        // --- Helpers ---
        private bool IsDefaultName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return true;
            if (name.StartsWith("パターン")) return true;
            if (name.StartsWith("追加パターン")) return true;
            if (name.StartsWith("ボタン")) return true;
            return false;
        }

        private string GenerateAutoName(PatternMacroConfig p)
        {
            if (p.TriggerType == 0) return $"ボタン{p.TriggerParam1}";
            if (p.TriggerType == 1) return $"ボタン{p.TriggerParam1}とボタン{p.TriggerParam2}";
            if (p.TriggerType == 2) return $"ボタン{p.TriggerParam1}を{p.TriggerParam2}回";
            return $"パターン";
        }

        private bool CheckDuplicate(PatternMacroConfig target)
        {
            foreach (var p in _patterns)
            {
                if (p == target) continue;
                if (p.TriggerType == target.TriggerType)
                {
                    if (target.TriggerType == 0 && target.TriggerParam1 == p.TriggerParam1) return true;
                    if (target.TriggerType == 1 && ((target.TriggerParam1 == p.TriggerParam1 && target.TriggerParam2 == p.TriggerParam2) || (target.TriggerParam1 == p.TriggerParam2 && target.TriggerParam2 == p.TriggerParam1))) return true;
                    if (target.TriggerType == 2 && target.TriggerParam1 == p.TriggerParam1 && target.TriggerParam2 == p.TriggerParam2) return true;
                }
            }
            return false;
        }

        private void RenderPatternCard(int index)
        {
            var pattern = _patterns[index];
            var card = new MaterialDesignThemes.Wpf.Card
            {
                Margin = new Thickness(0, 0, 0, 15),
                Padding = new Thickness(15)
            };

            var container = new StackPanel();

            var headerPanel = new Grid { Margin = new Thickness(0, 0, 0, 15) };
            headerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            
            Action triggerChanged = null;

            var headTitle = new TextBox
            {
                Text = string.IsNullOrWhiteSpace(pattern.Name) ? GenerateAutoName(pattern) : pattern.Name,
                FontWeight = FontWeights.Bold,
                FontSize = 16,
                Width = 200,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Left
            };
            headTitle.TextChanged += (s, e) => { pattern.Name = headTitle.Text; ScheduleAutoSync(pattern); };
            Grid.SetColumn(headTitle, 0);

            triggerChanged = () => {
                if (IsDefaultName(pattern.Name)) {
                    pattern.Name = GenerateAutoName(pattern);
                    headTitle.Text = pattern.Name;
                }
                ScheduleAutoSync(pattern);
            };

            // デフォルトパターンの判定：単押しで、パラメーターのボタン番号が現在の有効ボタン数以内であること
            // (並び替えによってindexが変化しても正しく判定できるようにする)
            bool isBasePattern = pattern.TriggerType == 0 && pattern.TriggerParam1 <= _activeButtonCount;

            var deleteBtn = new Button
            {
                Content = "✕",
                Foreground = new SolidColorBrush(Color.FromRgb(244, 67, 54)),
                ToolTip = "削除",
                FontSize = 18,
                Padding = new Thickness(5, 0, 5, 0),
                MinWidth = 30,
                Style = (Style)FindResource("MaterialDesignFlatButton"),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            Grid.SetColumn(deleteBtn, 1);

            if (isBasePattern)
            {
                deleteBtn.Visibility = Visibility.Collapsed;
            }

            deleteBtn.Click += (s, e) => 
            {
                if (MessageBox.Show("このパターンを削除しますか？", "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    _patterns.RemoveAt(index);
                    SavePatterns();
                    RenderAllPatterns();
                }
            };

            headerPanel.Children.Add(headTitle);
            headerPanel.Children.Add(deleteBtn);
            container.Children.Add(headerPanel);

            // Trigger Config Row
            var triggerPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 15) };
            
            var tTypeCombo = new ComboBox { Width = 150, Margin = new Thickness(0, 0, 15, 0), Style = (Style)FindResource("MaterialDesignFloatingHintComboBox") };
            MaterialDesignThemes.Wpf.HintAssist.SetHint(tTypeCombo, "トリガータイプ");
            if (isBasePattern)
            {
                tTypeCombo.Items.Add(new ComboBoxItem { Content = "単押し", Tag = 0 });
                tTypeCombo.SelectedIndex = 0;
                tTypeCombo.IsEnabled = false;
            }
            else
            {
                tTypeCombo.Items.Add(new ComboBoxItem { Content = "同時押し", Tag = 1 });
                tTypeCombo.Items.Add(new ComboBoxItem { Content = "複数回押し", Tag = 2 });
                if (pattern.TriggerType == 0) pattern.TriggerType = 1; // force migrate old patterns correctly
                if (pattern.TriggerType == 1) tTypeCombo.SelectedIndex = 0;
                else if (pattern.TriggerType == 2) tTypeCombo.SelectedIndex = 1;
                else tTypeCombo.SelectedIndex = 0;
            }

            var tParam1Combo = new ComboBox { Width = 100, Margin = new Thickness(0, 0, 15, 0), Style = (Style)FindResource("MaterialDesignFloatingHintComboBox") };
            MaterialDesignThemes.Wpf.HintAssist.SetHint(tParam1Combo, "ボタン");
            if (isBasePattern)
            {
                tParam1Combo.Items.Add(new ComboBoxItem { Content = $"ボタン{pattern.TriggerParam1}", Tag = pattern.TriggerParam1 });
                tParam1Combo.SelectedIndex = 0;
                tParam1Combo.IsEnabled = false;
            }
            else
            {
                for(int i=1; i<=_activeButtonCount; i++) tParam1Combo.Items.Add(new ComboBoxItem { Content = $"ボタン{i}", Tag = i });
                SetComboByTag(tParam1Combo, pattern.TriggerParam1);
                if (tParam1Combo.SelectedIndex < 0) tParam1Combo.SelectedIndex = 0;
            }

            var param2Panel = new StackPanel { Margin = new Thickness(0, 0, 15, 0) };
            var tParam2Combo = new ComboBox { Width = 100, Style = (Style)FindResource("MaterialDesignFloatingHintComboBox") };
            param2Panel.Children.Add(tParam2Combo);

            var repeatTxt = new TextBox { Width = 100, Text = pattern.RepeatInterval.ToString(), Style = (Style)FindResource("MaterialDesignFloatingHintTextBox") };
            MaterialDesignThemes.Wpf.HintAssist.SetHint(repeatTxt, "連続間隔(ms)");
            repeatTxt.TextChanged += (s, e) => { if (int.TryParse(repeatTxt.Text, out int v)) { pattern.RepeatInterval = v; ScheduleAutoSync(pattern); } };

            Action updateParam2UI = () =>
            {
                int pType = pattern.TriggerType;
                if (pType == 0)
                {
                    param2Panel.Visibility = Visibility.Collapsed;
                }
                else if (pType == 1)
                {
                    param2Panel.Visibility = Visibility.Visible;
                    MaterialDesignThemes.Wpf.HintAssist.SetHint(tParam2Combo, ""); // 同時押しの時の「ボタン2」ラベルを消す
                    int oldVal = pattern.TriggerParam2;
                    tParam2Combo.Items.Clear();
                    for (int i = 1; i <= _activeButtonCount; i++)
                    {
                        if (i != pattern.TriggerParam1) tParam2Combo.Items.Add(new ComboBoxItem { Content = $"ボタン{i}", Tag = i });
                    }
                    SetComboByTag(tParam2Combo, oldVal);
                    if (tParam2Combo.SelectedIndex < 0 && tParam2Combo.Items.Count > 0) tParam2Combo.SelectedIndex = 0;
                }
                else if (pType == 2)
                {
                    param2Panel.Visibility = Visibility.Visible;
                    MaterialDesignThemes.Wpf.HintAssist.SetHint(tParam2Combo, "回数");
                    int oldVal = pattern.TriggerParam2;
                    tParam2Combo.Items.Clear();
                    tParam2Combo.Items.Add(new ComboBoxItem { Content = "2回", Tag = 2 });
                    tParam2Combo.Items.Add(new ComboBoxItem { Content = "3回", Tag = 3 });
                    SetComboByTag(tParam2Combo, oldVal);
                    if (tParam2Combo.SelectedIndex < 0) tParam2Combo.SelectedIndex = 0;
                }
                if (tParam2Combo.SelectedItem is ComboBoxItem cbi) pattern.TriggerParam2 = (int)cbi.Tag;
            };

            updateParam2UI();

            bool isUpdating = false;

            tTypeCombo.SelectionChanged += (s, e) =>
            {
                if (isUpdating || tTypeCombo.SelectedItem == null) return;
                int oldType = pattern.TriggerType;
                int oldP2 = pattern.TriggerParam2;
                int newType = (int)((ComboBoxItem)tTypeCombo.SelectedItem).Tag;
                if (oldType == newType) return;

                pattern.TriggerType = newType;
                if (newType == 1 && pattern.TriggerParam1 == pattern.TriggerParam2) pattern.TriggerParam2 = pattern.TriggerParam1 == 1 ? 2 : 1;
                else if (newType == 2 && pattern.TriggerParam2 < 2) pattern.TriggerParam2 = 2;

                if (!isBasePattern && CheckDuplicate(pattern))
                {
                    MessageBox.Show("すでに同じ組み合わせのパターンが存在します。別の組み合わせを選択してください。", "重複エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    pattern.TriggerType = oldType;
                    pattern.TriggerParam2 = oldP2;
                    isUpdating = true;
                    SetComboByTag(tTypeCombo, oldType);
                    isUpdating = false;
                    return;
                }
                isUpdating = true;
                updateParam2UI();
                triggerChanged();
                isUpdating = false;
            };

            tParam1Combo.SelectionChanged += (s, e) =>
            {
                if (isUpdating || tParam1Combo.SelectedItem == null) return;
                int oldP1 = pattern.TriggerParam1;
                int newP1 = (int)((ComboBoxItem)tParam1Combo.SelectedItem).Tag;
                if (oldP1 == newP1) return;

                pattern.TriggerParam1 = newP1;
                int oldP2 = pattern.TriggerParam2;
                if (pattern.TriggerType == 1 && newP1 == pattern.TriggerParam2) pattern.TriggerParam2 = newP1 == 1 ? 2 : 1;

                if (!isBasePattern && CheckDuplicate(pattern))
                {
                    MessageBox.Show("すでに同じ組み合わせのパターンが存在します。別のボタンを選択してください。", "重複エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    pattern.TriggerParam1 = oldP1;
                    pattern.TriggerParam2 = oldP2;
                    isUpdating = true;
                    SetComboByTag(tParam1Combo, oldP1);
                    isUpdating = false;
                    return;
                }
                isUpdating = true;
                updateParam2UI();
                triggerChanged();
                isUpdating = false;
            };

            tParam2Combo.SelectionChanged += (s, e) =>
            {
                if (isUpdating || tParam2Combo.SelectedItem == null) return;
                int oldP2 = pattern.TriggerParam2;
                int newP2 = (int)((ComboBoxItem)tParam2Combo.SelectedItem).Tag;
                if (oldP2 == newP2) return;

                pattern.TriggerParam2 = newP2;

                if (!isBasePattern && CheckDuplicate(pattern))
                {
                    MessageBox.Show("すでに同じ組み合わせのパターンが存在します。別の値を選択してください。", "重複エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
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
            triggerPanel.Children.Add(param2Panel);
            triggerPanel.Children.Add(repeatTxt);
            container.Children.Add(triggerPanel);

            // For Mouse Capture Button logic
            var mouseCapMainBtn = new Button 
            { 
                Style = (Style)FindResource("MaterialDesignOutlinedButton"), 
                Margin = new Thickness(15, 0, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Left
            };

            Action updatePatternMouseBtnVisibility = () =>
            {
                bool hasMouse = pattern.Steps.Any(st => st.Type == "MOUSE");
                mouseCapMainBtn.Visibility = hasMouse ? Visibility.Visible : Visibility.Collapsed;
                
                if (_isCapturing && _capturingPattern == pattern) 
                {
                    mouseCapMainBtn.Content = _captureCount > 0 ? $"登録中 ({_captureCount})..." : "一括登録 停止 (Esc)";
                    mouseCapMainBtn.Foreground = new SolidColorBrush(Color.FromRgb(244, 67, 54));
                }
                else
                {
                    mouseCapMainBtn.Content = "マウス一括登録";
                    // テーマに合わせた文字色（ライト/ダーク両対応）
                    mouseCapMainBtn.Foreground = (Brush)FindResource("MaterialDesignBody");
                }
            };

            // ステップ用コンテナをGridで作成（D&Dのために行(Row)ベースで配置する）
            var stepsContainerGrid = new Grid();
            for (int i = 0; i < pattern.Steps.Count; i++)
            {
                stepsContainerGrid.RowDefinitions.Add(
                    new RowDefinition { Height = GridLength.Auto });
            }

            // Steps
            for (int i = 0; i < pattern.Steps.Count; i++)
            {
                int stepIndex = i;
                var step = pattern.Steps[stepIndex];

                var stepGrid = new Grid { Margin = new Thickness(0, 5, 0, 5) };
                stepGrid.Tag = step; // インデックスではなくオブジェクト参照自体を保持
                stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(28) }); // ドラッグハンドル
                stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });
                stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
                stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                // D&D用にGridの行番号を設定
                Grid.SetRow(stepGrid, stepIndex);

                // ドラッグハンドル（☰アイコン）
                var dragHandle = new TextBlock
                {
                    Text = "☰",
                    FontSize = 16,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Cursor = Cursors.Hand,
                    Foreground = (Brush)FindResource("MaterialDesignBodyLight"),
                    ToolTip = "ドラッグで並び替え"
                };
                Grid.SetColumn(dragHandle, 0);

                var translate = new TranslateTransform();
                stepGrid.RenderTransform = translate;
                
                bool isDragging = false;
                Point startMousePos = new Point();
                double itemHeight = 42.0;

                dragHandle.MouseLeftButtonDown += (s, e) =>
                {
                    isDragging = true;
                    translate.BeginAnimation(TranslateTransform.YProperty, null); // クリア
                    startMousePos = e.GetPosition(stepsContainerGrid);
                    dragHandle.CaptureMouse();
                    Panel.SetZIndex(stepGrid, 100);
                    stepGrid.Opacity = 0.8;
                };

                dragHandle.MouseMove += (s, e) =>
                {
                    if (!isDragging) return;
                    Point currentPos = e.GetPosition(stepsContainerGrid);
                    double offsetY = currentPos.Y - startMousePos.Y;
                    translate.Y = offsetY;

                    int currentIndex = pattern.Steps.IndexOf(step);
                    int stepsMoved = (int)Math.Round(offsetY / itemHeight);
                    int targetIndex = currentIndex + stepsMoved;
                    targetIndex = Math.Max(0, Math.Min(pattern.Steps.Count - 1, targetIndex));

                    if (targetIndex != currentIndex)
                    {
                        // リスト上の順序入れ替え
                        pattern.Steps.RemoveAt(currentIndex);
                        pattern.Steps.Insert(targetIndex, step);

                        // 各UI要素のGrid.Rowを更新して視覚的並び替え
                        foreach (UIElement child in stepsContainerGrid.Children)
                        {
                            if (child is Grid g && g.Tag is MacroStepConfig cfg)
                            {
                                int requiredRow = pattern.Steps.IndexOf(cfg);
                                int oldRow = Grid.GetRow(g);
                                if (oldRow != requiredRow)
                                {
                                    Grid.SetRow(g, requiredRow);
                                    if (g != stepGrid && g.RenderTransform is TranslateTransform targetTranslate)
                                    {
                                        targetTranslate.BeginAnimation(TranslateTransform.YProperty, null);
                                        targetTranslate.Y = (oldRow < requiredRow) ? -itemHeight : itemHeight;
                                        targetTranslate.BeginAnimation(
                                            TranslateTransform.YProperty,
                                            new DoubleAnimation(0, TimeSpan.FromMilliseconds(150))
                                            {
                                                EasingFunction = new System.Windows.Media.Animation.CircleEase
                                                {
                                                    EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut
                                                }
                                            });
                                    }
                                }
                            }
                        }
                        
                        startMousePos.Y += (targetIndex - currentIndex) * itemHeight;
                        translate.Y = currentPos.Y - startMousePos.Y;
                    }
                };

                dragHandle.MouseLeftButtonUp += (s, e) =>
                {
                    if (!isDragging) return;
                    isDragging = false;
                    dragHandle.ReleaseMouseCapture();
                    stepGrid.Opacity = 1.0;
                    Panel.SetZIndex(stepGrid, 0);

                    // 番号テキストの一括更新
                    foreach (UIElement child in stepsContainerGrid.Children)
                    {
                        if (child is Grid g && g.Tag is MacroStepConfig cfg)
                        {
                            int requiredRow = pattern.Steps.IndexOf(cfg);
                            var txt = g.Children.OfType<TextBlock>()
                                .FirstOrDefault(t => Grid.GetColumn(t) == 1);
                            if (txt != null) txt.Text = $"{requiredRow + 1}.";
                        }
                    }
                    
                    translate.BeginAnimation(
                        TranslateTransform.YProperty,
                        new DoubleAnimation(0, TimeSpan.FromMilliseconds(150))
                        {
                            EasingFunction = new System.Windows.Media.Animation.CircleEase
                            {
                                EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut
                            }
                        });

                    ScheduleAutoSync(pattern); // データ保存
                };

                var numBlock = new TextBlock
                {
                    Text = $"{stepIndex + 1}.",
                    VerticalAlignment = VerticalAlignment.Center,
                    Foreground = (Brush)FindResource("MaterialDesignBody")
                };
                Grid.SetColumn(numBlock, 1);

                var typeCombo = new ComboBox { Margin = new Thickness(0, 0, 10, 0), VerticalAlignment = VerticalAlignment.Center };
                typeCombo.Items.Add(new ComboBoxItem { Content = "キーボード", Tag = "KEY" });
                typeCombo.Items.Add(new ComboBoxItem { Content = "マウス座標", Tag = "MOUSE" });
                typeCombo.Items.Add(new ComboBoxItem { Content = "アプリ起動", Tag = "CMD" });
                typeCombo.Items.Add(new ComboBoxItem { Content = "待機", Tag = "WAIT" });
                SetComboByTag(typeCombo, step.Type);
                Grid.SetColumn(typeCombo, 2);

                var inputTxt = new TextBox
                {
                    Margin = new Thickness(0, 0, 10, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                    Text = step.Data,
                    Foreground = (Brush)FindResource("MaterialDesignBody")
                };
                Grid.SetColumn(inputTxt, 3);

                var browseBtn = new Button { Content = "参照", Margin = new Thickness(0, 0, 10, 0), VerticalAlignment = VerticalAlignment.Center };
                Grid.SetColumn(browseBtn, 4);
                
                // 削除ボタン（赤色はエラー強調のため維持）
                var stepDelBtn = new Button
                {
                    Content = "✕",
                    Style = (Style)FindResource("MaterialDesignFlatButton"),
                    Foreground = new SolidColorBrush(Color.FromRgb(244, 67, 54)),
                    FontSize = 18,
                    Padding = new Thickness(5, 0, 5, 0),
                    MinWidth = 30,
                    VerticalAlignment = VerticalAlignment.Center,
                    Visibility = stepIndex == 0 ? Visibility.Hidden : Visibility.Visible
                };
                Grid.SetColumn(stepDelBtn, 6);

                Action updateStepUI = () =>
                {
                    if (step.Type == "KEY") MaterialDesignThemes.Wpf.HintAssist.SetHint(inputTxt, "クリックして入力");
                    else if (step.Type == "MOUSE") MaterialDesignThemes.Wpf.HintAssist.SetHint(inputTxt, "座標 (例: 16000,8000)");
                    else if (step.Type == "CMD") MaterialDesignThemes.Wpf.HintAssist.SetHint(inputTxt, "コマンド/EXE");
                    else if (step.Type == "WAIT") MaterialDesignThemes.Wpf.HintAssist.SetHint(inputTxt, "待機ミリ秒 (例: 500)");

                    inputTxt.IsReadOnly = (step.Type == "KEY");
                    browseBtn.Visibility = (step.Type == "CMD") ? Visibility.Visible : Visibility.Collapsed;
                };

                typeCombo.SelectionChanged += (s, e) =>
                {
                    if (typeCombo.SelectedItem is ComboBoxItem c)
                    {
                        step.Type = c.Tag.ToString();
                        step.Data = "";
                        inputTxt.Text = "";
                        updateStepUI();
                        updatePatternMouseBtnVisibility();
                        ScheduleAutoSync(pattern);
                    }
                };

                inputTxt.TextChanged += (s, e) => { step.Data = inputTxt.Text; ScheduleAutoSync(pattern); };

                inputTxt.PreviewKeyDown += (s, e) =>
                {
                    if (step.Type == "KEY")
                    {
                        e.Handled = true;
                        var actualKey = e.Key == System.Windows.Input.Key.System ? e.SystemKey : e.Key;
                        if (actualKey == System.Windows.Input.Key.LeftCtrl || actualKey == System.Windows.Input.Key.RightCtrl ||
                            actualKey == System.Windows.Input.Key.LeftShift || actualKey == System.Windows.Input.Key.RightShift ||
                            actualKey == System.Windows.Input.Key.LeftAlt || actualKey == System.Windows.Input.Key.RightAlt ||
                            actualKey == System.Windows.Input.Key.LWin || actualKey == System.Windows.Input.Key.RWin) return;

                        string modifiers = "";
                        if (System.Windows.Input.Keyboard.Modifiers.HasFlag(System.Windows.Input.ModifierKeys.Control)) modifiers += "Ctrl+";
                        if (System.Windows.Input.Keyboard.Modifiers.HasFlag(System.Windows.Input.ModifierKeys.Shift)) modifiers += "Shift+";
                        if (System.Windows.Input.Keyboard.Modifiers.HasFlag(System.Windows.Input.ModifierKeys.Alt)) modifiers += "Alt+";

                        string keyStr = actualKey.ToString();
                        
                        // アルファベット単独キーの場合、Shiftなしなら小文字にする
                        if (actualKey >= System.Windows.Input.Key.A && actualKey <= System.Windows.Input.Key.Z)
                        {
                            if (!System.Windows.Input.Keyboard.Modifiers.HasFlag(System.Windows.Input.ModifierKeys.Shift))
                            {
                                keyStr = keyStr.ToLower();
                            }
                        }

                        // トップロー数字キーの対応 (D0-D9 -> 0-9)
                        if (actualKey >= System.Windows.Input.Key.D0 && actualKey <= System.Windows.Input.Key.D9)
                        {
                            keyStr = (actualKey - System.Windows.Input.Key.D0).ToString();
                        }

                        if (actualKey == System.Windows.Input.Key.Return) keyStr = "Enter";
                        if (actualKey == System.Windows.Input.Key.Escape) keyStr = "Esc";
                        if (actualKey == System.Windows.Input.Key.Space) keyStr = "Space";

                        inputTxt.Text = modifiers + keyStr;
                    }
                };

                browseBtn.Click += (s, e) =>
                {
                    var selector = new AppSelectorWindow();
                    selector.Owner = this;
                    if (selector.ShowDialog() == true && !string.IsNullOrEmpty(selector.SelectedExecutablePath))
                    {
                        inputTxt.Text = selector.SelectedExecutablePath;
                    }
                };

                stepDelBtn.Click += (s, e) =>
                {
                    if (stepIndex > 0)
                    {
                        pattern.Steps.RemoveAt(stepIndex);
                        RenderAllPatterns();
                    }
                };

                updateStepUI();

                stepGrid.Children.Add(dragHandle);
                stepGrid.Children.Add(numBlock);
                stepGrid.Children.Add(typeCombo);
                stepGrid.Children.Add(inputTxt);
                stepGrid.Children.Add(browseBtn);
                stepGrid.Children.Add(stepDelBtn);

                // ステップをGridコンテナに追加（StackPanelではなくGridに配置）
                stepsContainerGrid.Children.Add(stepGrid);
            }
            container.Children.Add(stepsContainerGrid);

            var footerPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 10, 0, 0) };

            var addStepBtn = new Button { Content = "＋ ステップ追加", Style = (Style)FindResource("MaterialDesignFlatButton"), HorizontalAlignment = HorizontalAlignment.Left };
            addStepBtn.Click += (s, e) =>
            {
                if (pattern.Steps.Count >= 10) MessageBox.Show("ステップ上限(10)です。");
                else { pattern.Steps.Add(new MacroStepConfig { Type = "KEY", Data = "" }); RenderAllPatterns(); }
            };
            if (pattern.Steps.Count < 10) footerPanel.Children.Add(addStepBtn);

            mouseCapMainBtn.Click += (s, e) =>
            {
                if (_isCapturing) StopCapture();
                else 
                {
                    int startIndex = pattern.Steps.Count;
                    for (int i=0; i<pattern.Steps.Count; i++) {
                        if (pattern.Steps[i].Type == "MOUSE" && string.IsNullOrEmpty(pattern.Steps[i].Data)) {
                            startIndex = i;
                            break;
                        }
                    }
                    StartCapture(pattern, startIndex);
                }
            };
            footerPanel.Children.Add(mouseCapMainBtn);

            updatePatternMouseBtnVisibility();

            container.Children.Add(footerPanel);

            card.Content = container;
            PatternsConfigPanel.Children.Add(card);
        }

        private void SetComboByTag(ComboBox cb, object tag)
        {
            string t = tag.ToString();
            for (int i = 0; i < cb.Items.Count; i++)
            {
                if ((cb.Items[i] as ComboBoxItem).Tag.ToString() == t) { cb.SelectedIndex = i; return; }
            }
        }

        // 自動送信をスケジュール（500ms後にまとめてSave+Sync）
        // 変更されたパターンオブジェクトの参照を保持し、送信時にTriggerTypeで判定する
        private void ScheduleAutoSync(PatternMacroConfig changedPattern = null)
        {
            if (_isRenderingPatterns) return;

            // 変更されたパターンの参照を保持
            if (changedPattern != null)
            {
                _unsyncedChangedButtons.Add(changedPattern.TriggerParam1 - 1);
                if (changedPattern.TriggerType == 1 && changedPattern.TriggerParam2 > 0)
                {
                    _unsyncedChangedButtons.Add(changedPattern.TriggerParam2 - 1);
                }
            }
            
            _autoSyncTimer.Stop();
            _autoSyncTimer.Start();
        }

        private void AddPatternButton_Click(object sender, RoutedEventArgs e)
        {
            if (_patterns.Count >= 30)
            {
                MessageBox.Show("パターンの登録上限に達しています。");
                return;
            }

            // 空いている設定を選ぶ（同時押し(1)優先、なければ複数回押し(2)）
            int tType = 1;
            int param1 = 1;
            int param2 = 2;
            bool foundVacant = false;

            // 同時押し(1)で空きを探す
            for (int i = 1; i <= _activeButtonCount; i++)
            {
                for (int j = 1; j <= _activeButtonCount; j++)
                {
                    if (i != j)
                    {
                        var temp = new PatternMacroConfig { TriggerType = 1, TriggerParam1 = i, TriggerParam2 = j };
                        if (!CheckDuplicate(temp))
                        {
                            tType = 1; param1 = i; param2 = j; foundVacant = true; break;
                        }
                    }
                }
                if (foundVacant) break;
            }

            // なければ複数回押し(2)で空きを探す
            if (!foundVacant)
            {
                for (int i = 1; i <= _activeButtonCount; i++)
                {
                    for (int j = 2; j <= 3; j++)
                    {
                        var temp = new PatternMacroConfig { TriggerType = 2, TriggerParam1 = i, TriggerParam2 = j };
                        if (!CheckDuplicate(temp))
                        {
                            tType = 2; param1 = i; param2 = j; foundVacant = true; break;
                        }
                    }
                    if (foundVacant) break;
                }
            }

            if (!foundVacant)
            {
                MessageBox.Show("全ての組み合わせが使用されています。既存のパターンを削除してから追加してください。");
                return;
            }

            var p = new PatternMacroConfig { TriggerType = tType, TriggerParam1 = param1, TriggerParam2 = param2 };
            p.Name = GenerateAutoName(p);
            p.Steps.Add(new MacroStepConfig { Type = "KEY", Data = "" });
            _patterns.Add(p);
            SavePatterns();
            RenderAllPatterns();
        }

        private void SyncAllButton_Click(object sender, RoutedEventArgs e)
        {
            SavePatterns();
            SyncAllToPico();
            
            if (_isConnected && _serialPort != null && _serialPort.IsOpen)
            {
                System.Threading.Thread.Sleep(300);
                try
                {
                    _serialPort.WriteLine("FLASH_ALL_BTNS");
                }
                catch { }
            }
            _unsyncedChangedButtons.Clear();
        }

        private void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_isConnected)
            {
                if (ComPortComboBox.SelectedItem != null)
                {
                    try
                    {
                        _serialPort = new SerialPort(ComPortComboBox.SelectedItem.ToString(), 115200);
                        // Arduinoからのシリアル通知を受信するイベントを登録
                        _serialPort.DataReceived += SerialPort_DataReceived;
                        _serialPort.Open();
                        _isConnected = true;
                        ConnectButton.Content = "切断する";
                        ComPortComboBox.IsEnabled = false;

                        try { File.WriteAllText(_comPortFilePath, _serialPort.PortName); } catch { }

                        // 接続成功したらデータ同期し、WAVEエフェクトを点灯させる
                        SyncAllToPico();
                        System.Threading.Thread.Sleep(50);
                        _serialPort.WriteLine("WAVE");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"接続失敗: {ex.Message}");
                    }
                }
            }
            else
            {
                if (_serialPort != null && _serialPort.IsOpen) _serialPort.Close();
                _isConnected = false;
                ConnectButton.Content = "接続する";
                ComPortComboBox.IsEnabled = true;
                // 切断時に連続動作状態をクリア
                _continuousActiveButtons.Clear();
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new SettingsWindow();
            settingsWindow.Owner = this;
            settingsWindow.ShowDialog();
            
            // Reload settings
            LoadSettings();
            RenderAllPatterns();
        }

        // --- Communication with Pico ---

        private void SyncAllToPico()
        {
            if (!_isConnected || _serialPort == null || !_serialPort.IsOpen) return;

            // 重複チェックはコンボボックス変更時に行うようにしたため、保存時は念のための簡略チェックのみでよい。
            for (int i = 0; i < _patterns.Count; i++)
            {
                for (int j = i + 1; j < _patterns.Count; j++)
                {
                    var p1 = _patterns[i];
                    var p2 = _patterns[j];
                    
                    bool isDuplicate = false;
                    if (p1.TriggerType == p2.TriggerType)
                    {
                        if (p1.TriggerType == 0 && p1.TriggerParam1 == p2.TriggerParam1) isDuplicate = true;
                        if (p1.TriggerType == 1 && ((p1.TriggerParam1 == p2.TriggerParam1 && p1.TriggerParam2 == p2.TriggerParam2) || (p1.TriggerParam1 == p2.TriggerParam2 && p1.TriggerParam2 == p2.TriggerParam1))) isDuplicate = true;
                        if (p1.TriggerType == 2 && p1.TriggerParam1 == p2.TriggerParam1 && p1.TriggerParam2 == p2.TriggerParam2) isDuplicate = true;
                    }

                    if (isDuplicate)
                    {
                        MessageBox.Show($"設定の重複があります。\n「{p1.Name}」と「{p2.Name}」が同じボタンの組み合わせに設定されています。\n修正してから再度保存・同期してください。", "重複エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }
            }

            try
            {
                _serialPort.WriteLine("CLEAR_ALL");
                System.Threading.Thread.Sleep(50);

                int limit = Math.Min(_patterns.Count, 30);
                for (int pIndex = 0; pIndex < limit; pIndex++)
                {
                    var p = _patterns[pIndex];
                    
                    // 空でない有用なステップだけを抽出
                    var validSteps = p.Steps.Where(st => !string.IsNullOrEmpty(st.Data) && st.Data != "NONE").ToList();
                    int steps = validSteps.Count;
                    
                    _serialPort.WriteLine($"ADD_PATTERN:{p.TriggerType}:{p.TriggerParam1}:{p.TriggerParam2}:{p.RepeatInterval}:{steps}");
                    System.Threading.Thread.Sleep(20);

                    for (int sIndex = 0; sIndex < steps; sIndex++)
                    {
                        var st = validSteps[sIndex];
                        string safeData = string.IsNullOrEmpty(st.Data) ? "NONE" : st.Data;
                        _serialPort.WriteLine($"SET_STEP:{pIndex}:{sIndex}:{st.Type}:{safeData}");
                        System.Threading.Thread.Sleep(20);
                    }
                }
                
                // すべて送信し終わってからEEPROMに一括保存させる
                _serialPort.WriteLine("SAVE_CONFIG");
                System.Threading.Thread.Sleep(50);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Sync error: " + ex.Message);
            }
        }

        // --- Global Hook Logic ---

        private void StartCapture(PatternMacroConfig pattern, int startStepIndex)
        {
            if (pattern.Steps.Count >= 10 && startStepIndex >= 10) return;

            _capturingPattern = pattern;
            _capturingStepIndex = startStepIndex;
            _isCapturing = true;
            _captureCount = 0;

            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                IntPtr handle = GetModuleHandle(curModule.ModuleName);
                _mouseHookID = SetWindowsHookEx(WH_MOUSE_LL, _mouseProc, handle, 0);
                _keyboardHookID = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardProc, handle, 0);
            }
            RenderAllPatterns();
        }

        private void StopCapture()
        {
            if (!_isCapturing) return;

            UnhookWindowsHookEx(_mouseHookID);
            UnhookWindowsHookEx(_keyboardHookID);
            _mouseHookID = IntPtr.Zero;
            _keyboardHookID = IntPtr.Zero;

            _isCapturing = false;

            foreach (var marker in _overlayMarkers) marker.Close();
            _overlayMarkers.Clear();

            _capturingPattern = null;
            _capturingStepIndex = -1;

            Application.Current.Dispatcher.Invoke(() => { RenderAllPatterns(); SavePatterns(); });
        }

        private void ShowClickMarker(int screenX, int screenY, int number)
        {
            PresentationSource source = PresentationSource.FromVisual(this);
            double dpiX = 1.0, dpiY = 1.0;
            if (source != null && source.CompositionTarget != null)
            {
                dpiX = source.CompositionTarget.TransformToDevice.M11;
                dpiY = source.CompositionTarget.TransformToDevice.M22;
            }
            double wpfX = screenX / dpiX;
            double wpfY = screenY / dpiY;

            string label = number <= 10 ? CircledNumbers[number - 1] : number.ToString();

            var markerWindow = new Window
            {
                WindowStyle = WindowStyle.None, AllowsTransparency = true, Background = Brushes.Transparent, Topmost = true, ShowInTaskbar = false, IsHitTestVisible = false, Width = 36, Height = 36, Left = wpfX - 18, Top = wpfY - 18, ResizeMode = ResizeMode.NoResize
            };

            var border = new Border { Background = new SolidColorBrush(Color.FromArgb(220, 255, 80, 80)), CornerRadius = new CornerRadius(18), Width = 36, Height = 36 };
            var text = new TextBlock { Text = label, Foreground = Brushes.White, FontSize = 18, FontWeight = FontWeights.Bold, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, TextAlignment = TextAlignment.Center };
            
            border.Child = text;
            markerWindow.Content = border;
            markerWindow.Show();
            _overlayMarkers.Add(markerWindow);
        }

        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONDOWN)
            {
                if (_isCapturing && _capturingPattern != null)
                {
                    MSLLHOOKSTRUCT hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

                    double screenWidth = SystemParameters.PrimaryScreenWidth;
                    double screenHeight = SystemParameters.PrimaryScreenHeight;
                    
                    PresentationSource source = PresentationSource.FromVisual(this);
                    double dpiX = 1.0, dpiY = 1.0;
                    if (source != null && source.CompositionTarget != null)
                    {
                        dpiX = source.CompositionTarget.TransformToDevice.M11;
                        dpiY = source.CompositionTarget.TransformToDevice.M22;
                    }

                    int physicalWidth = (int)(screenWidth * dpiX);
                    int physicalHeight = (int)(screenHeight * dpiY);

                    int hidX = (int)((hookStruct.pt.x / (double)physicalWidth) * 32767.0);
                    int hidY = (int)((hookStruct.pt.y / (double)physicalHeight) * 32767.0);
                    hidX = Math.Max(0, Math.Min(32767, hidX));
                    hidY = Math.Max(0, Math.Min(32767, hidY));
                    string data = $"{hidX},{hidY}";

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        var pattern = _capturingPattern;
                        _captureCount++;

                        ShowClickMarker(hookStruct.pt.x, hookStruct.pt.y, _captureCount);

                        if (_captureCount == 1 && _capturingStepIndex < pattern.Steps.Count)
                        {
                            pattern.Steps[_capturingStepIndex].Type = "MOUSE";
                            pattern.Steps[_capturingStepIndex].Data = data;
                        }
                        else if (pattern.Steps.Count < 10)
                        {
                            pattern.Steps.Add(new MacroStepConfig { Type = "MOUSE", Data = data });
                            if (pattern.Steps.Count >= 10) StopCapture();
                        }
                        else
                        {
                            StopCapture();
                        }
                        RenderAllPatterns(); // UI更新
                    });
                }
            }
            return CallNextHookEx(_mouseHookID, nCode, wParam, lParam);
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                if (vkCode == VK_ESCAPE)
                {
                    Application.Current.Dispatcher.Invoke(() => StopCapture());
                }
            }
            return CallNextHookEx(_keyboardHookID, nCode, wParam, lParam);
        }



        // =============================================
        // シリアル受信ハンドラ（Arduinoからの連続動作通知を処理）
        // =============================================
        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                if (_serialPort == null || !_serialPort.IsOpen) return;
                string line = _serialPort.ReadLine().Trim();

                if (line.StartsWith("CONTINUOUS_START:"))
                {
                    // 連続動作開始通知を処理
                    string[] parts = line.Substring(17).Split(':');
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        foreach (var part in parts)
                        {
                            if (int.TryParse(part, out int btnIdx))
                            {
                                _continuousActiveButtons.Add(btnIdx);
                            }
                        }
                    });
                }
                else if (line.StartsWith("CONTINUOUS_STOP:"))
                {
                    // 連続動作停止通知を処理
                    string btnStr = line.Substring(16);
                    if (int.TryParse(btnStr, out int btnIdx))
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            _continuousActiveButtons.Remove(btnIdx);
                        });
                    }
                }
            }
            catch { }
        }

        // =============================================
        // 連続動作中の警告表示
        // =============================================
        private void ShowContinuousWarning()
        {
            // 警告音を鳴らす（設定がONの場合）
            if (_warningSound)
            {
                SystemSounds.Exclamation.Play();
            }

            // 警告オーバーレイをフェードインで表示
            ContinuousWarningOverlay.Opacity = 0;
            ContinuousWarningOverlay.Visibility = Visibility.Visible;
            var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(200));
            ContinuousWarningOverlay.BeginAnimation(OpacityProperty, fadeIn);

            // 3秒後に自動で非表示にするタイマーを再起動
            _warningHideTimer.Stop();
            _warningHideTimer.Start();
        }

        // =============================================
        // 警告音設定の読み込み
        // =============================================
        private void LoadWarningSoundSetting()
        {
            if (File.Exists(_settingsFilePath))
            {
                try
                {
                    var json = JObject.Parse(File.ReadAllText(_settingsFilePath));
                    if (json["WarningSound"] != null)
                        _warningSound = json["WarningSound"].Value<bool>();
                }
                catch { }
            }
        }

        // 設定画面から呼び出される警告音設定の更新
        // 設定画面から呼び出される警告音設定の更新
        public void UpdateWarningSoundSetting(bool enabled)
        {
            _warningSound = enabled;
        }

        // =============================================
        // 自動アップデート確認 (GitHub API)
        // =============================================

        private const string GitHubOwner = "kazu-1234";
        private const string GitHubRepo = "LeftHandDevice";
        private const string GitHubApiUrl = "https://api.github.com/repos/{0}/{1}/releases/latest";

        /// <summary>
        /// 起動時にバックグラウンドでアップデートを確認する
        /// </summary>
        private async Task CheckUpdateAtStartupAsync()
        {
            try
            {
                // UIスレッドをブロックしないよう少し待機してから開始
                await Task.Delay(2000);

                using (var client = new HttpClient())
                {
                    // GitHub APIにはUser-Agentが必要
                    client.DefaultRequestHeaders.Add("User-Agent", "LeftHandDeviceApp");
                    var response = await client.GetAsync(string.Format(GitHubApiUrl, GitHubOwner, GitHubRepo));

                    if (response.IsSuccessStatusCode)
                    {
                        var json = await response.Content.ReadAsStringAsync();
                        var release = JObject.Parse(json);
                        string latestVersion = release["tag_name"]?.ToString().Trim().Replace("v", "");

                        if (!string.IsNullOrEmpty(latestVersion) && IsNewerVersion(latestVersion, SettingsWindow.AppVersion))
                        {
                            // 最新版がある場合、ダイアログを表示
                            UpdateDialogInfoText.Text = $"最新バージョン (v{latestVersion}) が利用可能です。\nアップデートして機能を最新に保ちましょう。";
                            MainDialogHost.IsOpen = true;
                        }
                    }
                }
            }
            catch { }
        }

        private bool IsNewerVersion(string latest, string current)
        {
            try
            {
                var vLatest = new Version(latest);
                var vCurrent = new Version(current);
                return vLatest > vCurrent;
            }
            catch { return false; }
        }

        /// <summary>
        /// ダイアログ内の「アップデートを行う」ボタンクリック
        /// </summary>
        private void UpdateDialog_GoClick(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = $"https://github.com/{GitHubOwner}/{GitHubRepo}/releases/latest",
                    UseShellExecute = true
                });
                MainDialogHost.IsOpen = false;
            }
            catch { }
        }
    }
}