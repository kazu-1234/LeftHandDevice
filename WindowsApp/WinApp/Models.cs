// Models.cs
using System.Collections.Generic;

namespace LeftHandDevice
{
    /// <summary>マクロ1ステップ（キー／マウス／コマンド／待機）。</summary>
    public class MacroStepConfig
    {
        public string Type { get; set; } = "KEY"; // KEY, MOUSE, CMD, WAIT
        public string Data { get; set; } = "";
    }

    /// <summary>トリガー条件付きパターン定義。</summary>
    public class PatternMacroConfig
    {
        public string Name { get; set; } = "";
        /// <summary>0=単押し, 1=同時押し, 2=複数回押し, 3=ボリューム</summary>
        public int TriggerType { get; set; } = 0;
        /// <summary>対象ボタン (1〜5)。ボリューム時はエンコーダ番号。</summary>
        public int TriggerParam1 { get; set; } = 1;
        /// <summary>同時押しのボタン2(1〜5)、または複数回押しの回数(2〜3)</summary>
        public int TriggerParam2 { get; set; } = 2;
        /// <summary>連続間隔 (ms)</summary>
        public int RepeatInterval { get; set; } = 200;
        public List<MacroStepConfig> Steps { get; set; } = new List<MacroStepConfig>();

        // ボリューム用の設定
        public int PotMin { get; set; } = 0;
        public int PotMax { get; set; } = 4095;
        public int VolLimit { get; set; } = 100;
    }
}
