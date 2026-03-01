using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace LeftHandDeviceApp
{
    public class MacroStepConfig
    {
        public string Type { get; set; } = "KEY"; // KEY, MOUSE, CMD, WAIT
        public string Data { get; set; } = "";
    }

    public class ButtonMacroConfig
    {
        public List<MacroStepConfig> Steps { get; set; } = new List<MacroStepConfig>();
    }

    public partial class MainWindow : Window
    {
        private SerialPort _serialPort;
        private bool _isConnected = false;
        private ButtonMacroConfig[] _macros = new ButtonMacroConfig[5];
        private StackPanel[] _buttonContainers = new StackPanel[5];

        private readonly string _comPortFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "saved_com_port.txt");

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
        private int _capturingButtonIndex = -1;
        private int _captureCount = 0;
        private List<Window> _overlayMarkers = new List<Window>();

        // 番号付き丸文字の配列
        private static readonly string[] CircledNumbers = {
            "❶", "❷", "❸", "❹", "❺",
            "❻", "❼", "❽", "❾", "❿"
        };

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
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        public MainWindow()
        {
            InitializeComponent();

            _mouseProc = MouseHookCallback;
            _keyboardProc = KeyboardHookCallback;

            for (int i = 0; i < 5; i++)
            {
                _macros[i] = new ButtonMacroConfig();
                // デフォルトとして1ステップ目をキーボード(A, B, C...)にする
                _macros[i].Steps.Add(new MacroStepConfig { Type = "KEY", Data = ((char)('a' + i)).ToString() });
            }

            LoadComPorts();
            GenerateBaseUI();

            // 自動接続
            if (ComPortComboBox.Items.Count > 0)
            {
                ConnectButton_Click(this, new RoutedEventArgs());
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            StopCapture();
            if (_serialPort != null && _serialPort.IsOpen)
            {
                _serialPort.Close();
            }
            base.OnClosed(e);
        }

        private void LoadComPorts()
        {
            ComPortComboBox.Items.Clear();
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                ComPortComboBox.Items.Add(port);
            }
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
                }
                else
                {
                    ComPortComboBox.SelectedIndex = 0;
                }
            }
        }

        private void GenerateBaseUI()
        {
            ButtonsConfigPanel.Children.Clear();
            for (int i = 0; i < 5; i++)
            {
                var card = new MaterialDesignThemes.Wpf.Card
                {
                    Margin = new Thickness(0, 0, 0, 15),
                    Padding = new Thickness(15)
                };

                var container = new StackPanel();
                _buttonContainers[i] = container;
                card.Content = container;
                ButtonsConfigPanel.Children.Add(card);

                RenderButtonConfig(i);
            }
        }

        private void RenderButtonConfig(int btnIndex)
        {
            var container = _buttonContainers[btnIndex];
            container.Children.Clear();

            // Header Row (Label + Actions)
            var headerGrid = new Grid { Margin = new Thickness(0, 0, 0, 10) };
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var title = new TextBlock
            {
                Text = $"ボタン {btnIndex + 1}",
                FontWeight = FontWeights.Bold,
                FontSize = 16,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(title, 0);

            var topActionsPanel = new StackPanel { Orientation = Orientation.Horizontal };
            Grid.SetColumn(topActionsPanel, 1);

            var captureBtn = new Button
            {
                Content = _isCapturing && _capturingButtonIndex == btnIndex ? 
                          (_captureCount > 0 ? $"登録中... ({_captureCount})" : "キャプチャ停止 (Esc)") : "マウス座標一括登録",
                Style = (Style)FindResource("MaterialDesignOutlinedButton"),
                Margin = new Thickness(0, 0, 10, 0)
            };
            if (_isCapturing && _capturingButtonIndex == btnIndex)
            {
                captureBtn.Foreground = Brushes.Red;
            }
            captureBtn.Click += (s, e) =>
            {
                if (_isCapturing) StopCapture();
                else StartCapture(btnIndex);
            };

            var clearBtn = new Button
            {
                Content = "全クリア",
                Style = (Style)FindResource("MaterialDesignOutlinedButton")
            };
            clearBtn.Click += (s, e) =>
            {
                _macros[btnIndex].Steps.Clear();
                _macros[btnIndex].Steps.Add(new MacroStepConfig { Type = "KEY", Data = "" });
                SyncButtonToPico(btnIndex);
                RenderButtonConfig(btnIndex);
            };

            topActionsPanel.Children.Add(captureBtn);
            topActionsPanel.Children.Add(clearBtn);
            headerGrid.Children.Add(title);
            headerGrid.Children.Add(topActionsPanel);
            container.Children.Add(headerGrid);

            // Steps
            var macro = _macros[btnIndex];
            for (int i = 0; i < macro.Steps.Count; i++)
            {
                int stepIndex = i; // local copy for closures
                var step = macro.Steps[stepIndex];

                var stepGrid = new Grid { Margin = new Thickness(0, 5, 0, 5) };
                stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });
                stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
                stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                stepGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                var numBlock = new TextBlock
                {
                    Text = $"{stepIndex + 1}.",
                    VerticalAlignment = VerticalAlignment.Center,
                    Foreground = (Brush)FindResource("MaterialDesignBody")
                };
                Grid.SetColumn(numBlock, 0);

                var typeCombo = new ComboBox
                {
                    Margin = new Thickness(0, 0, 10, 0),
                    VerticalAlignment = VerticalAlignment.Center
                };
                typeCombo.Items.Add(new ComboBoxItem { Content = "キーボード", Tag = "KEY" });
                typeCombo.Items.Add(new ComboBoxItem { Content = "マウス座標", Tag = "MOUSE" });
                typeCombo.Items.Add(new ComboBoxItem { Content = "アプリ起動", Tag = "CMD" });
                typeCombo.Items.Add(new ComboBoxItem { Content = "待機", Tag = "WAIT" });

                for (int j = 0; j < typeCombo.Items.Count; j++)
                {
                    if ((typeCombo.Items[j] as ComboBoxItem).Tag.ToString() == step.Type)
                    {
                        typeCombo.SelectedIndex = j;
                        break;
                    }
                }
                Grid.SetColumn(typeCombo, 1);

                var inputTextBox = new TextBox
                {
                    Margin = new Thickness(0, 0, 10, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                    Text = step.Data,
                    Foreground = (Brush)FindResource("MaterialDesignBody")
                };
                Grid.SetColumn(inputTextBox, 2);

                var browseBtn = new Button
                {
                    Content = "参照",
                    Margin = new Thickness(0, 0, 10, 0),
                    Visibility = step.Type == "CMD" ? Visibility.Visible : Visibility.Collapsed,
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(browseBtn, 3);
                
                browseBtn.Click += (s, e) =>
                {
                    var dlg = new OpenFileDialog { Filter = "実行可能ファイル (*.exe)|*.exe|すべて (*.*)|*.*" };
                    if (dlg.ShowDialog() == true) inputTextBox.Text = dlg.FileName;
                };

                var delBtn = new Button
                {
                    Content = "×",
                    Style = (Style)FindResource("MaterialDesignFlatButton"),
                    Foreground = Brushes.Red,
                    Padding = new Thickness(5, 0, 5, 0),
                    MinWidth = 30,
                    VerticalAlignment = VerticalAlignment.Center,
                    Visibility = stepIndex == 0 ? Visibility.Hidden : Visibility.Visible
                };
                Grid.SetColumn(delBtn, 4);

                delBtn.Click += (s, e) =>
                {
                    if (stepIndex == 0) return; // ステップ1は削除不可
                    _macros[btnIndex].Steps.RemoveAt(stepIndex);
                    SyncButtonToPico(btnIndex);
                    RenderButtonConfig(btnIndex);
                };

                // UI event bindings
                typeCombo.SelectionChanged += (s, e) =>
                {
                    if (typeCombo.SelectedItem is ComboBoxItem cbi)
                    {
                        string newType = cbi.Tag.ToString();
                        step.Type = newType;
                        step.Data = ""; // Clear data on type change
                        
                        if (newType == "KEY") MaterialDesignThemes.Wpf.HintAssist.SetHint(inputTextBox, "クリックして入力");
                        else if (newType == "MOUSE") MaterialDesignThemes.Wpf.HintAssist.SetHint(inputTextBox, "座標 (例: 16000,8000)");
                        else if (newType == "CMD") MaterialDesignThemes.Wpf.HintAssist.SetHint(inputTextBox, "コマンド (例: calc)");
                        else if (newType == "WAIT") MaterialDesignThemes.Wpf.HintAssist.SetHint(inputTextBox, "待機ミリ秒 (例: 500)");

                        inputTextBox.IsReadOnly = (newType == "KEY");
                        inputTextBox.Text = "";
                        browseBtn.Visibility = (newType == "CMD") ? Visibility.Visible : Visibility.Collapsed;
                        
                        SyncButtonToPico(btnIndex);
                    }
                };

                inputTextBox.PreviewKeyDown += (s, e) =>
                {
                    if (step.Type == "KEY")
                    {
                        e.Handled = true;
                        
                        // Handle System keys (when Alt is pressed, the main key becomes SystemKey)
                        var actualKey = e.Key == System.Windows.Input.Key.System ? e.SystemKey : e.Key;

                        // Ignore modifier keys alone
                        if (actualKey == System.Windows.Input.Key.LeftCtrl || actualKey == System.Windows.Input.Key.RightCtrl ||
                            actualKey == System.Windows.Input.Key.LeftShift || actualKey == System.Windows.Input.Key.RightShift ||
                            actualKey == System.Windows.Input.Key.LeftAlt || actualKey == System.Windows.Input.Key.RightAlt ||
                            actualKey == System.Windows.Input.Key.LWin || actualKey == System.Windows.Input.Key.RWin)
                            return;

                        string modifiers = "";
                        if (System.Windows.Input.Keyboard.Modifiers.HasFlag(System.Windows.Input.ModifierKeys.Control)) modifiers += "Ctrl+";
                        if (System.Windows.Input.Keyboard.Modifiers.HasFlag(System.Windows.Input.ModifierKeys.Shift)) modifiers += "Shift+";
                        if (System.Windows.Input.Keyboard.Modifiers.HasFlag(System.Windows.Input.ModifierKeys.Alt)) modifiers += "Alt+";

                        string keyStr = actualKey.ToString();
                        if (actualKey == System.Windows.Input.Key.Return) keyStr = "Enter";
                        if (actualKey == System.Windows.Input.Key.Escape) keyStr = "Esc";
                        if (actualKey == System.Windows.Input.Key.Space) keyStr = "Space";

                        inputTextBox.Text = modifiers + keyStr;
                    }
                };

                inputTextBox.TextChanged += (s, e) =>
                {
                    step.Data = inputTextBox.Text;
                    SyncButtonToPico(btnIndex);
                };

                stepGrid.Children.Add(numBlock);
                stepGrid.Children.Add(typeCombo);
                stepGrid.Children.Add(inputTextBox);
                stepGrid.Children.Add(browseBtn);
                stepGrid.Children.Add(delBtn);

                container.Children.Add(stepGrid);
            }

            var addBtn = new Button
            {
                Content = "＋ ステップ追加",
                Style = (Style)FindResource("MaterialDesignFlatButton"),
                Margin = new Thickness(0, 10, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Left
            };
            addBtn.Click += (s, e) =>
            {
                if (_macros[btnIndex].Steps.Count >= 10)
                {
                    MessageBox.Show("最大10ステップまでです。", "上限到達", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                _macros[btnIndex].Steps.Add(new MacroStepConfig { Type = "KEY", Data = "" });
                SyncButtonToPico(btnIndex);
                RenderButtonConfig(btnIndex);
            };

            if (macro.Steps.Count < 10)
            {
                container.Children.Add(addBtn);
            }
        }

        private void SyncButtonToPico(int btnIndex, bool triggerFlash = true)
        {
            if (!_isConnected || _serialPort == null || !_serialPort.IsOpen) return;

            var macro = _macros[btnIndex];
            
            try
            {
                _serialPort.WriteLine($"SETCOUNT:{btnIndex + 1}:{macro.Steps.Count}");
                System.Threading.Thread.Sleep(20); // 少し待つ

                for (int i = 0; i < macro.Steps.Count; i++)
                {
                    var step = macro.Steps[i];
                    string safeData = step.Data;
                    if (string.IsNullOrEmpty(safeData)) safeData = "0"; // To avoid empty parsing issues

                    _serialPort.WriteLine($"SET:{btnIndex + 1}:{i}:{step.Type}:{safeData}");
                    System.Threading.Thread.Sleep(20);
                }
                
                if (triggerFlash)
                {
                    System.Threading.Thread.Sleep(50); // FLASHコマンド前に少し余裕を持たせる
                    _serialPort.WriteLine($"FLASH:{btnIndex + 1}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Sync error: " + ex.Message);
            }
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
                        _serialPort.Open();
                        _isConnected = true;
                        ConnectButton.Content = "切断する";
                        ComPortComboBox.IsEnabled = false;

                        // On connect, sync all current UI state to Pico just in case (no flashy LED needed)
                        for(int i=0; i<5; i++) SyncButtonToPico(i, false);
                        
                        // 接続に成功したCOMポートを記憶しておく
                        try { File.WriteAllText(_comPortFilePath, _serialPort.PortName); } catch { }
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
            }
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new SettingsWindow();
            settingsWindow.Owner = this;
            settingsWindow.ShowDialog();
        }

        // --- Global Hook Logic ---
        private void StartCapture(int btnIndex)
        {
            if (_macros[btnIndex].Steps.Count >= 10)
            {
                MessageBox.Show("すでに10ステップ登録されています。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _capturingButtonIndex = btnIndex;
            _isCapturing = true;
            _captureCount = 0;

            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                IntPtr handle = GetModuleHandle(curModule.ModuleName);
                _mouseHookID = SetWindowsHookEx(WH_MOUSE_LL, _mouseProc, handle, 0);
                _keyboardHookID = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardProc, handle, 0);
            }

            // UI Refresh for specific button
            RenderButtonConfig(btnIndex);
        }

        private void StopCapture()
        {
            if (!_isCapturing) return;

            UnhookWindowsHookEx(_mouseHookID);
            UnhookWindowsHookEx(_keyboardHookID);
            _mouseHookID = IntPtr.Zero;
            _keyboardHookID = IntPtr.Zero;

            _isCapturing = false;

            // オーバーレイマーカーを全て閉じる
            foreach (var marker in _overlayMarkers)
            {
                marker.Close();
            }
            _overlayMarkers.Clear();

            int btnIndex = _capturingButtonIndex;
            _capturingButtonIndex = -1;

            if (btnIndex >= 0 && btnIndex < 5)
            {
                // キャプチャ完了時にまとめてPicoへ送信
                Application.Current.Dispatcher.Invoke(() =>
                {
                    SyncButtonToPico(btnIndex);
                    RenderButtonConfig(btnIndex);
                });
            }
        }

        // クリック位置に番号マーカーを表示する
        private void ShowClickMarker(int screenX, int screenY, int number)
        {
            // DPIを考慮してWPF座標に変換
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
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                Background = System.Windows.Media.Brushes.Transparent,
                Topmost = true,
                ShowInTaskbar = false,
                IsHitTestVisible = false,
                Width = 36,
                Height = 36,
                Left = wpfX - 18,
                Top = wpfY - 18,
                ResizeMode = ResizeMode.NoResize
            };

            var border = new System.Windows.Controls.Border
            {
                Background = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromArgb(220, 255, 80, 80)),
                CornerRadius = new CornerRadius(18),
                Width = 36,
                Height = 36
            };

            var text = new TextBlock
            {
                Text = label,
                Foreground = System.Windows.Media.Brushes.White,
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment = TextAlignment.Center
            };

            border.Child = text;
            markerWindow.Content = border;
            markerWindow.Show();
            _overlayMarkers.Add(markerWindow);
        }

        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONDOWN)
            {
                if (_isCapturing && _capturingButtonIndex != -1)
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
                        var macro = _macros[_capturingButtonIndex];
                        _captureCount++;

                        // クリック位置に番号マーカーを表示
                        ShowClickMarker(hookStruct.pt.x, hookStruct.pt.y, _captureCount);
                        
                        // データのみ更新（Picoへの送信はStopCaptureでまとめて行う）
                        if (macro.Steps.Count == 1 && string.IsNullOrEmpty(macro.Steps[0].Data))
                        {
                            macro.Steps[0].Type = "MOUSE";
                            macro.Steps[0].Data = data;
                        }
                        else if (macro.Steps.Count < 10)
                        {
                            macro.Steps.Add(new MacroStepConfig { Type = "MOUSE", Data = data });
                            if (macro.Steps.Count >= 10) StopCapture();
                        }
                        else
                        {
                            StopCapture();
                        }
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
    }
}