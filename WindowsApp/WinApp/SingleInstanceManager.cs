using System;
using System.Threading;

namespace LeftHandDevice
{
    /// <summary>
    /// 二重起動防止と、既存インスタンスへの表示／終了依頼（Mutex + 名前付きイベント）。
    /// </summary>
    internal static class SingleInstanceManager
    {
#if DEBUG
        private const string MutexName = "Global\\LeftHandDevice_SingleInstance_v1_DEBUG";
        private const string InteractiveShowEventName = "Global\\LeftHandDevice_ShowInteractive_v1_DEBUG";
        private const string ExitEventName = "Global\\LeftHandDevice_Exit_v1_DEBUG";
#else
        private const string MutexName = "Global\\LeftHandDevice_SingleInstance_v1";
        private const string InteractiveShowEventName = "Global\\LeftHandDevice_ShowInteractive_v1";
        private const string ExitEventName = "Global\\LeftHandDevice_Exit_v1";
#endif
        // 新しいプロジェクトへ複製する際は、他アプリと衝突しないよう上記の名前（AppName 部分）を変更すること。

        private static Mutex? _mutex;
        private static EventWaitHandle? _interactiveShowEvent;
        private static EventWaitHandle? _exitEvent;

        public static EventWaitHandle? InteractiveShowEvent => _interactiveShowEvent;
        public static EventWaitHandle? ExitEvent => _exitEvent;

        /// <param name="requestInteractiveShow">
        /// true のとき、既存インスタンスへ「ユーザー操作で GUI を開く」ことを通知する。
        /// --background の二重起動では false（通知しない）。
        /// </param>
        public static bool TryBecomePrimaryInstance(bool requestInteractiveShow)
        {
            _mutex = new Mutex(true, MutexName, out bool createdNew);
            if (!createdNew)
            {
                if (requestInteractiveShow)
                    SignalInteractiveShow();

                return false;
            }

            _interactiveShowEvent = new EventWaitHandle(
                false,
                EventResetMode.AutoReset,
                InteractiveShowEventName);
            _exitEvent = new EventWaitHandle(
                false,
                EventResetMode.AutoReset,
                ExitEventName);

            return true;
        }

        public static void SignalInteractiveShow()
        {
            try
            {
                using var showEvent = EventWaitHandle.OpenExisting(InteractiveShowEventName);
                showEvent.Set();
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                // 既存インスタンスが応答しない場合は無視する
            }
        }

        /// <summary>既存インスタンスへ終了を依頼する（インストーラ / --exit 用）。</summary>
        public static void SignalExit()
        {
            try
            {
                using var exitEvent = EventWaitHandle.OpenExisting(ExitEventName);
                exitEvent.Set();
            }
            catch (WaitHandleCannotBeOpenedException)
            {
            }
        }

        public static void Release()
        {
            _interactiveShowEvent?.Dispose();
            _interactiveShowEvent = null;
            _exitEvent?.Dispose();
            _exitEvent = null;

            if (_mutex != null)
            {
                try { _mutex.ReleaseMutex(); } catch { }
                _mutex.Dispose();
                _mutex = null;
            }
        }
    }
}
