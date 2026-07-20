// v2.0.0
using System;

namespace LeftHandDevice
{
    /// <summary>ページ間で共有するアプリ状態。</summary>
    public sealed class AppState
    {
        public AppState(Settings settings, DeviceService device)
        {
            Settings = settings;
            Device = device;
        }

        public Settings Settings { get; }
        public DeviceService Device { get; }
    }
}
