// DeviceService.cs
// プロセス寿命のデバイス／パターン／シリアル／ボリューム制御（UI非依存）
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Media;
using Microsoft.UI.Dispatching;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace LeftHandDevice
{
    /// <summary>VolumeController から呼び出されるホスト側コールバック。</summary>
    public interface IVolumeHost
    {
        void OnExternalVolumeChanged();
        void ClearVolumeHintTitle();
    }

    /// <summary>
    /// シリアル接続・パターン保存／同期・ボリューム制御を担うバックエンド。
    /// AppRuntime が所有し、プロセス寿命で生き続ける。
    /// </summary>
    public sealed class DeviceService : IVolumeHost, IDisposable
    {
        private readonly DispatcherQueue _dispatcherQueue;
        private readonly string _comPortFilePath;
        private readonly string _settingsFilePath;
        private readonly string _patternsFilePath;

        private SerialPort? _serialPort;
        private bool _isConnected;
        private VolumeController? _volumeController;

        private List<PatternMacroConfig> _patterns = new();
        private int _activeButtonCount = 5;
        private bool _warningSound = true;
        public int ActiveVolumeCount { get; set; } = 1;

        private PatternMacroConfig? _calibrationTarget;
        private int _calibrationTargetProperty; // 1=PotMin, 2=PotMax
        private Action<int>? _calibrationTargetAction;

        private readonly HashSet<int> _continuousActiveButtons = new();
        private readonly HashSet<int> _unsyncedChangedButtons = new();
        private DateTime _lastPotNeedZeroHint = DateTime.MinValue;

        private DispatcherQueueTimer? _autoSyncTimer;
        private bool _suppressAutoSync;
        private bool _disposed;
        private readonly object _serialGate = new();

        public event Action? PatternsChanged;
        public event Action? ConnectionChanged;
        /// <summary>連続動作中の警告表示を UI に依頼する。</summary>
        public event Action? ContinuousWarningRequested;
        /// <summary>null でタイトルヒント解除、非 null でヒント文言。</summary>
        public event Action<string?>? VolumeHintChanged;
        /// <summary>(volumeIndex, adcValue)</summary>
        public event Action<int, int>? AdcReceived;

        public bool IsConnected =>
            _isConnected && _serialPort != null && _serialPort.IsOpen;

        public int ActiveButtonCount => _activeButtonCount;
        public bool WarningSound => _warningSound;

        public IReadOnlyList<PatternMacroConfig> Patterns => _patterns;
        public IReadOnlyCollection<int> ContinuousActiveButtons => _continuousActiveButtons;

        public DeviceService(DispatcherQueue dispatcherQueue)
        {
            _dispatcherQueue = dispatcherQueue
                ?? throw new ArgumentNullException(nameof(dispatcherQueue));

            string baseDir = Path.GetDirectoryName(Environment.ProcessPath)
                ?? AppDomain.CurrentDomain.BaseDirectory;

            _comPortFilePath = Path.Combine(baseDir, "saved_com_port.txt");
            _settingsFilePath = Path.Combine(baseDir, "app_settings.json");
            _patternsFilePath = Path.Combine(baseDir, "app_patterns.json");

            _volumeController = new VolumeController(OnExternalVolumeChanged, _dispatcherQueue);

            _autoSyncTimer = _dispatcherQueue.CreateTimer();
            _autoSyncTimer.IsRepeating = false;
            _autoSyncTimer.Interval = TimeSpan.FromMilliseconds(200);
            _autoSyncTimer.Tick += AutoSyncTimer_Tick;

            LoadDeviceSettings();
            LoadPatterns();
            ActiveVolumeCount = 1;
        }

        // ---------- ポート ----------

        /// <summary>
        /// 利用可能な COM ポート一覧。
        /// GetPortNames は環境によって時間がかかるため、UI スレッドから直接呼ばず Task.Run 推奨。
        /// </summary>
        public string[] GetAvailablePorts()
        {
            try
            {
                string[] ports = SerialPort.GetPortNames();
                Array.Sort(ports, StringComparer.OrdinalIgnoreCase);
                return ports;
            }
            catch
            {
                return Array.Empty<string>();
            }
        }

        public string? LoadSavedComPort()
        {
            try
            {
                if (File.Exists(_comPortFilePath))
                    return File.ReadAllText(_comPortFilePath).Trim();
            }
            catch { }
            return null;
        }

        /// <summary>
        /// シリアル接続。同期・Sleep を含むため必ずバックグラウンドで呼ぶこと。
        /// VolumeController.Start は UI スレッド向けに別途 StartVolumeMonitoring を呼ぶ。
        /// </summary>
        public bool Connect(string port)
        {
            if (string.IsNullOrWhiteSpace(port) || _disposed)
                return false;

            lock (_serialGate)
            {
                if (IsConnected)
                    DisconnectCore(sendVolumeModeOff: true, raiseEvents: false, waitForClose: true);

                try
                {
                    var sp = new SerialPort(port, 115200)
                    {
                        NewLine = "\n",
                        ReadTimeout = 500,
                        WriteTimeout = 2000,
                        DtrEnable = true,
                        RtsEnable = true
                    };
                    sp.DataReceived += SerialPort_DataReceived;
                    sp.Open();
                    _serialPort = sp;
                    _isConnected = true;

                    try { File.WriteAllText(_comPortFilePath, sp.PortName); } catch { }

                    SyncAllToPico();
                    System.Threading.Thread.Sleep(50);
                    // 音量は常時 HID。残留した PC モードをクリアする
                    SendSerialCommand("PC_VOLUME_MODE:0");
                    SendSerialCommand("VOL_RESET");
                    try { sp.WriteLine("WAVE"); } catch { }

                    RaiseConnectionChanged();
                    return true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Connect failed: " + ex.Message);
                    CleanupSerialPort(waitForClose: false);
                    _isConnected = false;
                    RaiseConnectionChanged();
                    return false;
                }
            }
        }

        /// <summary>接続成功後に UI スレッドから呼ぶ（NAudio 初期化）。</summary>
        public void StartVolumeMonitoring()
        {
            if (_disposed || !IsConnected) return;
            try { _volumeController?.Start(); } catch { }
        }

        public void Disconnect()
        {
            lock (_serialGate)
            {
                DisconnectCore(sendVolumeModeOff: true, raiseEvents: true, waitForClose: true);
            }
        }

        private void DisconnectCore(bool sendVolumeModeOff, bool raiseEvents, bool waitForClose)
        {
            if (sendVolumeModeOff && !_disposed && _serialPort != null && _serialPort.IsOpen)
            {
                try { SendSerialCommand("PC_VOLUME_MODE:0"); } catch { }
            }

            try { _volumeController?.Stop(); } catch { }
            try { ClearVolumeHintTitle(); } catch { }
            CleanupSerialPort(waitForClose);
            _isConnected = false;
            _continuousActiveButtons.Clear();
            if (raiseEvents)
            {
                try { RaiseConnectionChanged(); } catch { }
            }
        }

        /// <summary>
        /// SerialPort.Close は DataReceived/ReadLine 待ちで固まりやすいため BG で閉じる。
        /// waitForClose 時も短時間のみ待ち、UI フリーズを避ける。
        /// </summary>
        private void CleanupSerialPort(bool waitForClose)
        {
            SerialPort? port = _serialPort;
            _serialPort = null;
            if (port == null) return;

            try { port.DataReceived -= SerialPort_DataReceived; } catch { }

            var closeTask = System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    if (port.IsOpen)
                    {
                        try { port.DiscardInBuffer(); } catch { }
                        try { port.DiscardOutBuffer(); } catch { }
                        port.Close();
                    }
                }
                catch { }
                try { port.Dispose(); } catch { }
            });

            if (!waitForClose)
                return;

            try { closeTask.Wait(400); } catch { }
        }

        // ---------- 設定（exe 隣 app_settings.json） ----------

        private void LoadDeviceSettings()
        {
            if (!File.Exists(_settingsFilePath)) return;
            try
            {
                var json = JObject.Parse(File.ReadAllText(_settingsFilePath));
                if (json["ActiveButtonCount"] != null)
                    _activeButtonCount = json["ActiveButtonCount"]!.Value<int>();
                if (json["WarningSound"] != null)
                    _warningSound = json["WarningSound"]!.Value<bool>();
            }
            catch { }
            ActiveVolumeCount = 1;
        }

        /// <summary>theme / LastUpdateCheck など既存キーを保持したまま部分更新する。</summary>
        private void PatchAppSettings(Action<JObject> patch)
        {
            try
            {
                JObject settings;
                if (File.Exists(_settingsFilePath))
                    settings = JObject.Parse(File.ReadAllText(_settingsFilePath));
                else
                    settings = new JObject();

                patch(settings);
                File.WriteAllText(_settingsFilePath, settings.ToString());
            }
            catch { }
        }

        public void SetActiveButtonCount(int count)
        {
            count = Math.Clamp(count, 1, 5);
            if (_activeButtonCount == count) return;
            _activeButtonCount = count;
            PatchAppSettings(s => s["ActiveButtonCount"] = count);
            EnsureBasePatterns();
            PatternListHelper.EnsureVolumePattern(_patterns);
            RaisePatternsChanged();
        }

        public void SetWarningSound(bool enabled)
        {
            _warningSound = enabled;
            PatchAppSettings(s => s["WarningSound"] = enabled);
        }

        public void SetLastUpdateCheck(string displayText)
        {
            PatchAppSettings(s => s["LastUpdateCheck"] = displayText);
        }

        public string? GetLastUpdateCheck()
        {
            if (!File.Exists(_settingsFilePath)) return null;
            try
            {
                var json = JObject.Parse(File.ReadAllText(_settingsFilePath));
                return json["LastUpdateCheck"]?.ToString();
            }
            catch { return null; }
        }

        // ---------- パターン ----------

        public void LoadPatterns()
        {
            _suppressAutoSync = true;
            try
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

                PatternListHelper.MigrateVolumePatterns(_patterns);

                // 初回起動時など空の場合、不足分を補う
                if (!File.Exists(_patternsFilePath) && _patterns.Count < 5)
                {
                    var existingBtnIds = _patterns
                        .Where(p => p.TriggerType == 0)
                        .Select(p => p.TriggerParam1)
                        .ToList();
                    for (int i = 1; i <= 5; i++)
                    {
                        if (!existingBtnIds.Contains(i))
                        {
                            var p = new PatternMacroConfig
                            {
                                TriggerType = 0,
                                TriggerParam1 = i,
                                Name = $"ボタン{i}"
                            };
                            p.Steps.Add(new MacroStepConfig
                            {
                                Type = "KEY",
                                Data = ((char)('a' + (i - 1))).ToString()
                            });
                            _patterns.Add(p);
                        }
                    }
                    _patterns = _patterns.OrderBy(p => p.TriggerParam1).ToList();
                    SavePatterns();
                }

                EnsureBasePatterns();
                PatternListHelper.EnsureVolumePattern(_patterns);
            }
            finally
            {
                _suppressAutoSync = false;
            }

            RaisePatternsChanged();
        }

        public void ReloadPatternsFromDisk()
        {
            LoadDeviceSettings();
            LoadPatterns();
        }

        public void SavePatterns()
        {
            try
            {
                string json = JsonConvert.SerializeObject(_patterns, Formatting.Indented);
                File.WriteAllText(_patternsFilePath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SavePatterns failed: " + ex.Message);
            }
        }

        /// <summary>有効ボタン数の範囲内で単押しベースパターンを補完する。</summary>
        private void EnsureBasePatterns()
        {
            for (int i = 1; i <= _activeButtonCount; i++)
            {
                if (_patterns.Any(p => p.TriggerType == 0 && p.TriggerParam1 == i))
                    continue;

                var mutated = _patterns.FirstOrDefault(p => p.Name == $"ボタン{i}");
                if (mutated != null)
                {
                    mutated.TriggerType = 0;
                    mutated.TriggerParam1 = i;
                }
                else
                {
                    var newBase = new PatternMacroConfig
                    {
                        TriggerType = 0,
                        TriggerParam1 = i,
                        Name = $"ボタン{i}"
                    };
                    newBase.Steps.Add(new MacroStepConfig { Type = "KEY", Data = "" });
                    _patterns.Insert(Math.Min(i - 1, _patterns.Count), newBase);
                }
            }
        }

        /// <summary>
        /// 空きトリガーを探してパターンを追加する。
        /// 失敗時は null（上限 or 組み合わせ尽きた）。
        /// </summary>
        public PatternMacroConfig? AddPattern()
        {
            if (_patterns.Count >= 30)
                return null;

            int tType = 1;
            int param1 = 1;
            int param2 = 2;
            bool foundVacant = false;

            for (int i = 1; i <= _activeButtonCount; i++)
            {
                for (int j = 1; j <= _activeButtonCount; j++)
                {
                    if (i == j) continue;
                    var temp = new PatternMacroConfig
                    {
                        TriggerType = 1,
                        TriggerParam1 = i,
                        TriggerParam2 = j
                    };
                    if (!CheckDuplicate(temp))
                    {
                        tType = 1;
                        param1 = i;
                        param2 = j;
                        foundVacant = true;
                        break;
                    }
                }
                if (foundVacant) break;
            }

            if (!foundVacant)
            {
                for (int i = 1; i <= _activeButtonCount; i++)
                {
                    for (int j = 2; j <= 3; j++)
                    {
                        var temp = new PatternMacroConfig
                        {
                            TriggerType = 2,
                            TriggerParam1 = i,
                            TriggerParam2 = j
                        };
                        if (!CheckDuplicate(temp))
                        {
                            tType = 2;
                            param1 = i;
                            param2 = j;
                            foundVacant = true;
                            break;
                        }
                    }
                    if (foundVacant) break;
                }
            }

            if (!foundVacant)
                return null;

            var p = new PatternMacroConfig
            {
                TriggerType = tType,
                TriggerParam1 = param1,
                TriggerParam2 = param2
            };
            p.Name = GenerateAutoName(p);
            p.Steps.Add(new MacroStepConfig { Type = "KEY", Data = "" });
            _patterns.Add(p);
            SavePatterns();
            RaisePatternsChanged();
            return p;
        }

        public bool DeletePattern(PatternMacroConfig pattern)
        {
            if (!_patterns.Remove(pattern))
                return false;
            SavePatterns();
            ScheduleAutoSync(null);
            RaisePatternsChanged();
            return true;
        }

        public bool DeletePatternAt(int index)
        {
            if (index < 0 || index >= _patterns.Count)
                return false;
            _patterns.RemoveAt(index);
            SavePatterns();
            ScheduleAutoSync(null);
            RaisePatternsChanged();
            return true;
        }

        /// <summary>fromIndex の要素を toIndex へ移動する。</summary>
        public bool ReorderPatterns(int fromIndex, int toIndex, bool scheduleSync = true, bool raiseChanged = true)
        {
            if (fromIndex < 0 || fromIndex >= _patterns.Count) return false;
            if (toIndex < 0 || toIndex >= _patterns.Count) return false;
            if (fromIndex == toIndex) return true;

            var item = _patterns[fromIndex];
            _patterns.RemoveAt(fromIndex);
            _patterns.Insert(toIndex, item);
            SavePatterns();
            if (scheduleSync)
                ScheduleAutoSync(null);
            if (raiseChanged)
                RaisePatternsChanged();
            return true;
        }

        /// <summary>新しい順序でリストを置き換える（同一インスタンス集合であること）。</summary>
        /// <param name="scheduleSync">true のときマイコン同期をスケジュール（UI フリーズ防止のため DnD では false）</param>
        /// <param name="raiseChanged">false のとき UI 再描画イベントを抑止（ブロック DnD ドロップ後のチラつき防止）</param>
        public void ReorderPatterns(IList<PatternMacroConfig> newOrder, bool scheduleSync = true, bool raiseChanged = true)
        {
            if (newOrder == null || newOrder.Count != _patterns.Count)
                throw new ArgumentException("newOrder must contain the same patterns.");

            _patterns = new List<PatternMacroConfig>(newOrder);
            SavePatterns();
            if (scheduleSync)
                ScheduleAutoSync(null);
            if (raiseChanged)
                RaisePatternsChanged();
        }

        // ---------- ヘルパー（公開） ----------

        public bool IsDefaultName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return true;
            if (name.StartsWith("パターン")) return true;
            if (name.StartsWith("追加パターン")) return true;
            if (name.StartsWith("ボタン")) return true;
            return false;
        }

        public string GenerateAutoName(PatternMacroConfig p)
        {
            if (p.TriggerType == 0) return $"ボタン{p.TriggerParam1}";
            if (p.TriggerType == 1) return $"ボタン{p.TriggerParam1}とボタン{p.TriggerParam2}";
            if (p.TriggerType == 2) return $"ボタン{p.TriggerParam1}を{p.TriggerParam2}回";
            if (p.TriggerType == 3) return "ボリューム（エンコーダ）";
            return "パターン";
        }

        public bool CheckDuplicate(PatternMacroConfig target)
        {
            foreach (var p in _patterns)
            {
                if (p == target) continue;
                if (p.TriggerType != target.TriggerType) continue;

                if (target.TriggerType == 0 && target.TriggerParam1 == p.TriggerParam1)
                    return true;
                if (target.TriggerType == 1 &&
                    ((target.TriggerParam1 == p.TriggerParam1 && target.TriggerParam2 == p.TriggerParam2) ||
                     (target.TriggerParam1 == p.TriggerParam2 && target.TriggerParam2 == p.TriggerParam1)))
                    return true;
                if (target.TriggerType == 2 &&
                    target.TriggerParam1 == p.TriggerParam1 &&
                    target.TriggerParam2 == p.TriggerParam2)
                    return true;
                if (target.TriggerType == 3 && target.TriggerParam1 == p.TriggerParam1)
                    return true;
            }
            return false;
        }

        // ---------- 自動同期 ----------

        public void ScheduleAutoSync(PatternMacroConfig? changedPattern = null)
        {
            if (_suppressAutoSync) return;
            if (_autoSyncTimer == null) return;

            if (changedPattern != null)
            {
                _unsyncedChangedButtons.Add(changedPattern.TriggerParam1 - 1);
                if (changedPattern.TriggerType == 1 && changedPattern.TriggerParam2 > 0)
                    _unsyncedChangedButtons.Add(changedPattern.TriggerParam2 - 1);
            }

            _autoSyncTimer.Stop();
            _autoSyncTimer.Start();
        }

        private void AutoSyncTimer_Tick(DispatcherQueueTimer sender, object args)
        {
            sender.Stop();
            SavePatterns();

            // Sleep 付きのシリアル送信は UI スレッドでやるとフリーズするためバックグラウンドへ
            var buttons = new HashSet<int>(_unsyncedChangedButtons);
            _unsyncedChangedButtons.Clear();

            _ = System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    SyncAllToPico();

                    if (buttons.Count > 0 && IsConnected && _serialPort != null)
                    {
                        System.Threading.Thread.Sleep(300);
                        foreach (int bIndex in buttons)
                        {
                            if (bIndex >= 0 && bIndex < 5)
                            {
                                try
                                {
                                    _serialPort.WriteLine($"FLASH_BUTTONS:{bIndex}:-1");
                                    System.Threading.Thread.Sleep(350);
                                }
                                catch { }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("AutoSync background: " + ex.Message);
                }
            });
        }

        // ---------- Pico 同期 ----------

        /// <summary>
        /// CLEAR_ALL / ADD_PATTERN / SET_STEP / SET_POT_CONFIG / SAVE_CONFIG を送信。
        /// 重複がある場合は false。
        /// </summary>
        public bool SyncAllToPico()
        {
            if (_disposed || !IsConnected || _serialPort == null)
                return false;

            for (int i = 0; i < _patterns.Count; i++)
            {
                for (int j = i + 1; j < _patterns.Count; j++)
                {
                    var p1 = _patterns[i];
                    var p2 = _patterns[j];

                    bool isDuplicate = false;
                    if (p1.TriggerType == p2.TriggerType)
                    {
                        if (p1.TriggerType == 0 && p1.TriggerParam1 == p2.TriggerParam1)
                            isDuplicate = true;
                        if (p1.TriggerType == 1 &&
                            ((p1.TriggerParam1 == p2.TriggerParam1 && p1.TriggerParam2 == p2.TriggerParam2) ||
                             (p1.TriggerParam1 == p2.TriggerParam2 && p1.TriggerParam2 == p2.TriggerParam1)))
                            isDuplicate = true;
                        if (p1.TriggerType == 2 &&
                            p1.TriggerParam1 == p2.TriggerParam1 &&
                            p1.TriggerParam2 == p2.TriggerParam2)
                            isDuplicate = true;
                        if (p1.TriggerType == 3 && p1.TriggerParam1 == p2.TriggerParam1)
                            isDuplicate = true;
                    }

                    if (isDuplicate)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"Sync duplicate: '{p1.Name}' / '{p2.Name}'");
                        return false;
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
                    var validSteps = p.Steps
                        .Where(st => !string.IsNullOrEmpty(st.Data) && st.Data != "NONE")
                        .ToList();
                    int steps = validSteps.Count;

                    _serialPort.WriteLine(
                        $"ADD_PATTERN:{p.TriggerType}:{p.TriggerParam1}:{p.TriggerParam2}:{p.RepeatInterval}:{steps}");
                    System.Threading.Thread.Sleep(20);

                    for (int sIndex = 0; sIndex < steps; sIndex++)
                    {
                        var st = validSteps[sIndex];
                        string safeData = string.IsNullOrEmpty(st.Data) ? "NONE" : st.Data;
                        _serialPort.WriteLine($"SET_STEP:{pIndex}:{sIndex}:{st.Type}:{safeData}");
                        System.Threading.Thread.Sleep(20);
                    }
                }

                SyncPotConfig();
                _serialPort.WriteLine("SAVE_CONFIG");
                System.Threading.Thread.Sleep(50);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Sync error: " + ex.Message);
                return false;
            }
        }

        /// <summary>手動「すべて同期」用。同期後に FLASH_ALL_BTNS。</summary>
        public bool SyncAllToPicoAndFlash()
        {
            SavePatterns();
            bool ok = SyncAllToPico();
            if (ok && IsConnected && _serialPort != null)
            {
                System.Threading.Thread.Sleep(300);
                try { _serialPort.WriteLine("FLASH_ALL_BTNS"); } catch { }
            }
            _unsyncedChangedButtons.Clear();
            return ok;
        }

        public void SendSerialCommand(string cmd)
        {
            if (_disposed || !IsConnected || _serialPort == null) return;
            try { _serialPort.WriteLine(cmd); } catch { }
        }

        public void SyncPotConfig()
        {
            var vp = _patterns.FirstOrDefault(p => p.TriggerType == 3 && p.TriggerParam1 == 1);
            if (vp == null) return;
            vp.PotMin = 0;
            vp.PotMax = 4095;
            vp.VolLimit = 100;
            SendSerialCommand("SET_POT_CONFIG:0:4095:100");
            System.Threading.Thread.Sleep(50);
            SendSerialCommand("SAVE_CONFIG");
        }

        public PatternMacroConfig? GetVolumePattern()
            => PatternListHelper.GetVolumePattern(_patterns);

        /// <summary>PoT 校正を開始。接続中のみ true。</summary>
        public bool BeginVolumeCalibration(PatternMacroConfig vol, int property, Action<int> onSampled)
        {
            if (!IsConnected || _serialPort == null)
                return false;

            _calibrationTarget = vol;
            _calibrationTargetProperty = property;
            _calibrationTargetAction = onSampled;
            SendSerialCommand("GET_ADC");
            return true;
        }

        // ---------- IVolumeHost ----------

        public void OnExternalVolumeChanged()
        {
            ClearVolumeHintTitle();
        }

        public void ClearVolumeHintTitle()
        {
            VolumeHintChanged?.Invoke(null);
        }

        private void ShowPotNeedZeroHint()
        {
            var now = DateTime.UtcNow;
            if ((now - _lastPotNeedZeroHint).TotalMilliseconds < 1500) return;
            _lastPotNeedZeroHint = now;
            VolumeHintChanged?.Invoke("つまみを最小位置へ戻してください");
        }

        /// <summary>
        /// UI（ウィンドウ非アクティブ等）から呼ぶ。連続動作中なら警告音＋イベント発火。
        /// </summary>
        public void RequestContinuousWarning()
        {
            if (_continuousActiveButtons.Count == 0) return;
            if (_warningSound)
            {
                try { SystemSounds.Exclamation.Play(); } catch { }
            }
            ContinuousWarningRequested?.Invoke();
        }

        // ---------- シリアル受信 ----------

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                if (_serialPort == null || !_serialPort.IsOpen) return;
                string line = _serialPort.ReadLine().Trim();

                if (line.StartsWith("ADC:"))
                {
                    var parts = line.Substring(4).Split(':');
                    if (parts.Length == 2)
                    {
                        if (int.TryParse(parts[0], out int volIdx) && int.TryParse(parts[1], out int val))
                        {
                            _dispatcherQueue.TryEnqueue(() => HandleAdcSample(volIdx, val));
                        }
                    }
                    else if (parts.Length == 1)
                    {
                        if (int.TryParse(parts[0], out int val))
                        {
                            _dispatcherQueue.TryEnqueue(() => HandleAdcSample(1, val));
                        }
                    }
                }
                else if (line.StartsWith("VOL_STEP:"))
                {
                    string dirStr = line.Substring(9).Trim();
                    if (int.TryParse(dirStr, out int dir) && dir != 0)
                    {
                        _dispatcherQueue.TryEnqueue(() =>
                        {
                            _volumeController?.ApplyDeviceVolumeStep(dir);
                            ClearVolumeHintTitle();
                        });
                    }
                }
                else if (line == "POT_NEED_ZERO")
                {
                    _dispatcherQueue.TryEnqueue(ShowPotNeedZeroHint);
                }
                else if (line.StartsWith("CONTINUOUS_START:"))
                {
                    string[] parts = line.Substring(17).Split(':');
                    _dispatcherQueue.TryEnqueue(() =>
                    {
                        foreach (var part in parts)
                        {
                            if (int.TryParse(part, out int btnIdx))
                                _continuousActiveButtons.Add(btnIdx);
                        }
                    });
                }
                else if (line.StartsWith("CONTINUOUS_STOP:"))
                {
                    string btnStr = line.Substring(16);
                    if (int.TryParse(btnStr, out int btnIdx))
                    {
                        _dispatcherQueue.TryEnqueue(() =>
                        {
                            _continuousActiveButtons.Remove(btnIdx);
                        });
                    }
                }
            }
            catch { }
        }

        private void HandleAdcSample(int volIdx, int val)
        {
            if (_calibrationTarget != null && _calibrationTargetAction != null)
            {
                if (_calibrationTargetProperty == 1) _calibrationTarget.PotMin = val;
                else if (_calibrationTargetProperty == 2) _calibrationTarget.PotMax = val;
                PatternListHelper.NormalizePotEndpoints(_calibrationTarget);

                var action = _calibrationTargetAction;
                _calibrationTarget = null;
                _calibrationTargetProperty = 0;
                _calibrationTargetAction = null;

                action(val);
                SavePatterns();
                SyncPotConfig();
            }

            AdcReceived?.Invoke(volIdx, val);
        }

        private void RaisePatternsChanged() => PatternsChanged?.Invoke();
        private void RaiseConnectionChanged() => ConnectionChanged?.Invoke();

        // ---------- Dispose ----------

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _suppressAutoSync = true;

            if (_autoSyncTimer != null)
            {
                _autoSyncTimer.Stop();
                _autoSyncTimer.Tick -= AutoSyncTimer_Tick;
                _autoSyncTimer = null;
            }

            try
            {
                // 終了時も PC_VOLUME_MODE:0 を送り HID に戻す（Close 待ち・イベントなし）
                lock (_serialGate)
                {
                    DisconnectCore(sendVolumeModeOff: true, raiseEvents: false, waitForClose: false);
                }
            }
            catch { }

            try { _volumeController?.Dispose(); } catch { }
            _volumeController = null;
        }
    }
}
