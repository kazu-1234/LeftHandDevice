// PatternListHelper.cs
using System.Collections.Generic;
using System.Linq;

namespace LeftHandDevice
{
    public static class PatternListHelper
    {
        public static void MigrateVolumePatterns(List<PatternMacroConfig> patterns)
        {
            var volList = patterns.Where(p => p.TriggerType == 3).ToList();
            if (volList.Count == 0) return;

            var keep = volList.FirstOrDefault(p => p.TriggerParam2 == 1)
                ?? volList.FirstOrDefault(p => p.Name != null && p.Name.Contains("右"))
                ?? volList[0];

            foreach (var p in volList)
            {
                if (p.PotMin != 0 && keep.PotMin == 0) keep.PotMin = p.PotMin;
                if (p.PotMax != 4095 && keep.PotMax == 4095) keep.PotMax = p.PotMax;
                if (p.VolLimit != 100 && keep.VolLimit == 100) keep.VolLimit = p.VolLimit;
                if (p.Steps.Count > 0 && keep.Steps.Count == 0)
                    keep.Steps = new List<MacroStepConfig>(p.Steps);
            }

            keep.TriggerType = 3;
            keep.TriggerParam1 = 1;
            keep.TriggerParam2 = 0;
            keep.Name = "ボリューム";
            PreservePotEndpoints(keep);

            foreach (var p in volList)
            {
                if (p != keep) patterns.Remove(p);
            }
        }

        public static void EnsureVolumePattern(List<PatternMacroConfig> patterns)
        {
            MigrateVolumePatterns(patterns);
            if (!patterns.Any(p => p.TriggerType == 3 && p.TriggerParam1 == 1))
            {
                patterns.Add(new PatternMacroConfig
                {
                    TriggerType = 3,
                    TriggerParam1 = 1,
                    TriggerParam2 = 0,
                    Name = "ボリューム",
                    VolLimit = 100,
                    PotMin = 0,
                    PotMax = 4095
                });
            }
        }

        public static PatternMacroConfig? GetVolumePattern(List<PatternMacroConfig> patterns)
        {
            return patterns.FirstOrDefault(p => p.TriggerType == 3 && p.TriggerParam1 == 1);
        }

        public static void PreservePotEndpoints(PatternMacroConfig vol)
        {
            // PotMinは「ユーザーが0%として登録した物理位置」、
            // PotMaxは「100%として登録した物理位置」。大小順に入れ替えると回転方向が壊れる。
            _ = vol;
        }

        public static void NormalizePotEndpoints(PatternMacroConfig vol) => PreservePotEndpoints(vol);

        /// <summary>校正済みPoT位置を0〜100%で表示</summary>
        public static int AdcToVolumePercent(int adc, int potMin, int potMax)
        {
            int range = potMax - potMin;
            if (range == 0) return 0;

            int percent = (adc - potMin) * 100 / range;
            return System.Math.Clamp(percent, 0, 100);
        }

        /// <summary>ファームへ送る1段あたりの最大ステップ数（+2%刻み）</summary>
        public static int MaxVolumeSteps(int volLimitPercent)
        {
            int steps = volLimitPercent / 2;
            return steps < 1 ? 1 : steps;
        }
    }
}
