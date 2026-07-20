// VolumeController.cs
using System;
using System.Threading.Tasks;
using Microsoft.UI.Dispatching;
using NAudio.CoreAudioApi;

namespace LeftHandDevice
{
    /// <summary>
    /// マスター音量の監視と、デバイスエンコーダからの ±2% ステップ適用。
    /// WPF 非依存。UI スレッドへのマーシャリングは DispatcherQueue 経由。
    /// </summary>
    public sealed class VolumeController : IDisposable
    {
        private readonly Action _onExternalVolumeChanged;
        private readonly DispatcherQueue? _dispatcherQueue;
        private MMDevice? _device;
        private bool _changeFromApp;
        private bool _suppressExternalOnce;
        private DateTime _lastExternalNotify = DateTime.MinValue;
        private const int ExternalNotifyCooldownMs = 400;

        public bool IsActive => _device != null;

        /// <param name="onExternalVolumeChanged">外部（キーボード等）からの音量変更時コールバック</param>
        /// <param name="dispatcherQueue">指定時はコールバックを UI スレッドへマーシャリング</param>
        public VolumeController(Action onExternalVolumeChanged, DispatcherQueue? dispatcherQueue = null)
        {
            _onExternalVolumeChanged = onExternalVolumeChanged
                ?? throw new ArgumentNullException(nameof(onExternalVolumeChanged));
            _dispatcherQueue = dispatcherQueue;
        }

        public void Start()
        {
            Stop();
            try
            {
                var enumerator = new MMDeviceEnumerator();
                _device = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                _device.AudioEndpointVolume.OnVolumeNotification += OnVolumeNotification;
                _suppressExternalOnce = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("VolumeController.Start: " + ex.Message);
                _device = null;
            }
        }

        public void Stop()
        {
            if (_device != null)
            {
                try
                {
                    _device.AudioEndpointVolume.OnVolumeNotification -= OnVolumeNotification;
                }
                catch { }
                _device.Dispose();
                _device = null;
            }
        }

        public void Dispose() => Stop();

        /// <summary>ファームから VOL_STEP:±1 を受信したとき ±2%（0〜1スケールで0.02）</summary>
        public void ApplyDeviceVolumeStep(int direction)
        {
            if (_device == null || direction == 0) return;

            try
            {
                _changeFromApp = true;
                float current = _device.AudioEndpointVolume.MasterVolumeLevelScalar;
                float delta = direction > 0 ? 0.02f : -0.02f;
                float next = Math.Clamp(current + delta, 0.0f, 1.0f);
                _device.AudioEndpointVolume.MasterVolumeLevelScalar = next;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ApplyDeviceVolumeStep: " + ex.Message);
            }
            finally
            {
                Task.Delay(80).ContinueWith(_ => _changeFromApp = false);
            }
        }

        private void OnVolumeNotification(AudioVolumeNotificationData data)
        {
            if (_changeFromApp || _suppressExternalOnce)
            {
                if (_suppressExternalOnce)
                    _suppressExternalOnce = false;
                return;
            }

            var now = DateTime.UtcNow;
            if ((now - _lastExternalNotify).TotalMilliseconds < ExternalNotifyCooldownMs)
                return;
            _lastExternalNotify = now;

            if (_dispatcherQueue != null)
            {
                _dispatcherQueue.TryEnqueue(() => _onExternalVolumeChanged());
            }
            else
            {
                _onExternalVolumeChanged();
            }
        }
    }
}
