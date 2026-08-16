using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1.PresentationHost
{
    /// <summary>换页过渡 effect+dir ↔ PpEntryEffect（对外只暴露字符串）。</summary>
    internal static class PptTransitionEffectMap
    {
        // 与 Office15 PIA 数值对齐（Wpp 晚绑定共用）
        public const int EffectNone = 0;
        public const int EffectCut = 257;
        public const int EffectCoverLeft = 1281;
        public const int EffectCoverUp = 1282;
        public const int EffectCoverRight = 1283;
        public const int EffectCoverDown = 1284;
        public const int EffectFade = 1793;
        public const int EffectUncoverLeft = 2049;
        public const int EffectUncoverUp = 2050;
        public const int EffectUncoverRight = 2051;
        public const int EffectUncoverDown = 2052;
        public const int EffectWipeLeft = 2817;
        public const int EffectWipeUp = 2818;
        public const int EffectWipeRight = 2819;
        public const int EffectWipeDown = 2820;
        public const int EffectFadeSmoothly = 3849;
        public const int EffectPushDown = 3852;
        public const int EffectPushLeft = 3853;
        public const int EffectPushRight = 3854;
        public const int EffectPushUp = 3855;

        private static readonly HashSet<string> EffectsNeedingDir = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "push", "wipe", "cover", "uncover"
        };

        public static bool NeedsDir(string effect)
        {
            return !string.IsNullOrWhiteSpace(effect) && EffectsNeedingDir.Contains(effect.Trim());
        }

        public static bool TryResolveToEntryEffect(
            string effect,
            string dir,
            out int entryEffect,
            out string error)
        {
            entryEffect = EffectNone;
            error = null;
            if (string.IsNullOrWhiteSpace(effect))
            {
                error = "须提供 effect";
                return false;
            }

            string e = effect.Trim().ToLowerInvariant();
            string d = string.IsNullOrWhiteSpace(dir) ? null : dir.Trim().ToLowerInvariant();

            if (NeedsDir(e))
            {
                if (string.IsNullOrEmpty(d))
                {
                    error = "effect=" + e + " 须提供 dir（left|right|up|down）";
                    return false;
                }

                if (d != "left" && d != "right" && d != "up" && d != "down")
                {
                    error = "非法 dir: " + dir + "（须为 left|right|up|down）";
                    return false;
                }
            }
            else if (!string.IsNullOrEmpty(d))
            {
                // 不需要 dir 时忽略多余 dir，避免模型误传导致失败
            }

            switch (e)
            {
                case "none":
                    entryEffect = EffectNone;
                    return true;
                case "cut":
                    entryEffect = EffectCut;
                    return true;
                case "fade":
                    entryEffect = EffectFade;
                    return true;
                case "push":
                    entryEffect = DirToPush(d);
                    return true;
                case "wipe":
                    entryEffect = DirToWipe(d);
                    return true;
                case "cover":
                    entryEffect = DirToCover(d);
                    return true;
                case "uncover":
                    entryEffect = DirToUncover(d);
                    return true;
                default:
                    error = "不支持的 effect: " + effect
                        + "（白名单：none|fade|push|wipe|cover|uncover|cut）";
                    return false;
            }
        }

        public static void TryDescribe(
            int entryEffect,
            out string effect,
            out string dir,
            out string warning)
        {
            effect = null;
            dir = null;
            warning = null;

            switch (entryEffect)
            {
                case EffectNone:
                    effect = "none";
                    return;
                case EffectCut:
                    effect = "cut";
                    return;
                case EffectFade:
                case EffectFadeSmoothly:
                    effect = "fade";
                    return;
                case EffectPushLeft:
                    effect = "push";
                    dir = "left";
                    return;
                case EffectPushRight:
                    effect = "push";
                    dir = "right";
                    return;
                case EffectPushUp:
                    effect = "push";
                    dir = "up";
                    return;
                case EffectPushDown:
                    effect = "push";
                    dir = "down";
                    return;
                case EffectWipeLeft:
                    effect = "wipe";
                    dir = "left";
                    return;
                case EffectWipeRight:
                    effect = "wipe";
                    dir = "right";
                    return;
                case EffectWipeUp:
                    effect = "wipe";
                    dir = "up";
                    return;
                case EffectWipeDown:
                    effect = "wipe";
                    dir = "down";
                    return;
                case EffectCoverLeft:
                    effect = "cover";
                    dir = "left";
                    return;
                case EffectCoverRight:
                    effect = "cover";
                    dir = "right";
                    return;
                case EffectCoverUp:
                    effect = "cover";
                    dir = "up";
                    return;
                case EffectCoverDown:
                    effect = "cover";
                    dir = "down";
                    return;
                case EffectUncoverLeft:
                    effect = "uncover";
                    dir = "left";
                    return;
                case EffectUncoverRight:
                    effect = "uncover";
                    dir = "right";
                    return;
                case EffectUncoverUp:
                    effect = "uncover";
                    dir = "up";
                    return;
                case EffectUncoverDown:
                    effect = "uncover";
                    dir = "down";
                    return;
                default:
                    effect = "unknown";
                    warning = "未识别的 EntryEffect=" + entryEffect;
                    return;
            }
        }

        public static PowerPoint.PpEntryEffect ToPowerPoint(int entryEffect)
        {
            return (PowerPoint.PpEntryEffect)entryEffect;
        }

        private static int DirToPush(string d)
        {
            switch (d)
            {
                case "left":
                    return EffectPushLeft;
                case "right":
                    return EffectPushRight;
                case "up":
                    return EffectPushUp;
                default:
                    return EffectPushDown;
            }
        }

        private static int DirToWipe(string d)
        {
            switch (d)
            {
                case "left":
                    return EffectWipeLeft;
                case "right":
                    return EffectWipeRight;
                case "up":
                    return EffectWipeUp;
                default:
                    return EffectWipeDown;
            }
        }

        private static int DirToCover(string d)
        {
            switch (d)
            {
                case "left":
                    return EffectCoverLeft;
                case "right":
                    return EffectCoverRight;
                case "up":
                    return EffectCoverUp;
                default:
                    return EffectCoverDown;
            }
        }

        private static int DirToUncover(string d)
        {
            switch (d)
            {
                case "left":
                    return EffectUncoverLeft;
                case "right":
                    return EffectUncoverRight;
                case "up":
                    return EffectUncoverUp;
                default:
                    return EffectUncoverDown;
            }
        }
    }
}
