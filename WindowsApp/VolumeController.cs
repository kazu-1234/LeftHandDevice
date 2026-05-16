// VolumeController.cs
// v1.17.0
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using NAudio.CoreAudioApi;

namespace LeftHandDeviceApp
{
    /// <summary>
    /// マスター音量の監視と、デバイスつまみからの +2% ステップ適用。
    /// 外部（キーボード等）で音量が変わったらファームへ VOL_RESET を送る。
    /// </summary>
    public sealed class VolumeController : IDisposable
    {
        private readonly MainWindow _main;
        private MMDevice? _device;
        private bool _changeFromApp;
        private bool _suppressExternalOnce;
        private DateTime _lastExternalNotify = DateTime.MinValue;
        private const int ExternalNotifyCooldownMs = 400;

        public bool IsActive => _device != null;

        public VolumeController(MainWindow main)
        {
            _main = main;
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

        /// <summary>ファームから VOL_STEP:1 を受信したとき +2%（0〜1スケールで0.02）</summary>
        public void ApplyDeviceVolumeStep()
        {
            if (_device == null) return;

            try
            {
                _changeFromApp = true;
                float current = _device.AudioEndpointVolume.MasterVolumeLevelScalar;
                float next = Math.Min(1.0f, current + 0.02f);
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

            Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
            {
                _main.OnExternalVolumeChanged();
            }));
        }
    }
}
