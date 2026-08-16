using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace WordAddIn1.PresentationHost
{
    /// <summary>首期动画白名单 ↔ COM 映射（可删减，禁止默默扩展）。</summary>
    internal static class PptAnimationEffectMap
    {
        public static bool TryResolve(
            string category,
            string effect,
            string dir,
            out PowerPoint.MsoAnimEffect effectId,
            out bool isExit,
            out bool needsDir,
            out PowerPoint.MsoAnimDirection direction,
            out string error)
        {
            effectId = PowerPoint.MsoAnimEffect.msoAnimEffectAppear;
            isExit = false;
            needsDir = false;
            direction = PowerPoint.MsoAnimDirection.msoAnimDirectionLeft;
            error = null;

            string cat = (category ?? "").Trim().ToLowerInvariant();
            string eff = (effect ?? "").Trim().ToLowerInvariant();
            if (cat != "entrance" && cat != "emphasis" && cat != "exit")
            {
                error = "category 须为 entrance|emphasis|exit";
                return false;
            }

            isExit = cat == "exit";

            switch (eff)
            {
                case "appear":
                    if (cat == "emphasis")
                    {
                        error = "appear 不适用于 emphasis";
                        return false;
                    }

                    effectId = PowerPoint.MsoAnimEffect.msoAnimEffectAppear;
                    return true;
                case "fade":
                    effectId = PowerPoint.MsoAnimEffect.msoAnimEffectFade;
                    return true;
                case "fly_in":
                    if (cat != "entrance")
                    {
                        error = "fly_in 仅用于 entrance";
                        return false;
                    }

                    needsDir = true;
                    effectId = PowerPoint.MsoAnimEffect.msoAnimEffectFly;
                    return TryParseDir(dir, out direction, out error);
                case "fly_out":
                    if (cat != "exit")
                    {
                        error = "fly_out 仅用于 exit";
                        return false;
                    }

                    needsDir = true;
                    isExit = true;
                    effectId = PowerPoint.MsoAnimEffect.msoAnimEffectFly;
                    return TryParseDir(dir, out direction, out error);
                case "wipe":
                    if (cat == "emphasis")
                    {
                        error = "wipe 不适用于 emphasis";
                        return false;
                    }

                    needsDir = true;
                    effectId = PowerPoint.MsoAnimEffect.msoAnimEffectWipe;
                    return TryParseDir(dir, out direction, out error);
                case "float_up":
                    if (cat != "entrance")
                    {
                        error = "float_up 仅用于 entrance";
                        return false;
                    }

                    effectId = PowerPoint.MsoAnimEffect.msoAnimEffectAscend;
                    return true;
                case "grow_shrink":
                    if (cat != "emphasis")
                    {
                        error = "grow_shrink 仅用于 emphasis";
                        return false;
                    }

                    effectId = PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink;
                    return true;
                case "spin":
                    if (cat != "emphasis")
                    {
                        error = "spin 仅用于 emphasis";
                        return false;
                    }

                    effectId = PowerPoint.MsoAnimEffect.msoAnimEffectSpin;
                    return true;
                default:
                    error = "不支持的 effect: " + eff;
                    return false;
            }
        }

        public static bool TryParseTrigger(string trigger, out PowerPoint.MsoAnimTriggerType triggerType, out string error)
        {
            triggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
            error = null;
            string t = string.IsNullOrWhiteSpace(trigger) ? "on_click" : trigger.Trim().ToLowerInvariant();
            switch (t)
            {
                case "on_click":
                    triggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                    return true;
                case "with_previous":
                    triggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                    return true;
                case "after_previous":
                    triggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;
                    return true;
                default:
                    error = "trigger 须为 on_click|with_previous|after_previous";
                    return false;
            }
        }

        public static string TriggerToName(PowerPoint.MsoAnimTriggerType trigger)
        {
            switch (trigger)
            {
                case PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious:
                    return "with_previous";
                case PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious:
                    return "after_previous";
                default:
                    return "on_click";
            }
        }

        public static void DescribeEffect(
            PowerPoint.Effect effect,
            out string category,
            out string effectName,
            out string dir)
        {
            category = "entrance";
            effectName = "appear";
            dir = null;

            bool isExit = false;
            try
            {
                isExit = effect.Exit == Office.MsoTriState.msoTrue
                    || (int)effect.Exit == -1;
            }
            catch (Exception)
            {
            }

            PowerPoint.MsoAnimEffect type = PowerPoint.MsoAnimEffect.msoAnimEffectAppear;
            try
            {
                type = effect.EffectType;
            }
            catch (Exception)
            {
            }

            switch (type)
            {
                case PowerPoint.MsoAnimEffect.msoAnimEffectAppear:
                    effectName = "appear";
                    category = isExit ? "exit" : "entrance";
                    break;
                case PowerPoint.MsoAnimEffect.msoAnimEffectFade:
                    effectName = "fade";
                    category = isExit ? "exit" : "entrance";
                    break;
                case PowerPoint.MsoAnimEffect.msoAnimEffectFly:
                    effectName = isExit ? "fly_out" : "fly_in";
                    category = isExit ? "exit" : "entrance";
                    dir = DirToName(TryReadDirection(effect));
                    break;
                case PowerPoint.MsoAnimEffect.msoAnimEffectWipe:
                    effectName = "wipe";
                    category = isExit ? "exit" : "entrance";
                    dir = DirToName(TryReadDirection(effect));
                    break;
                case PowerPoint.MsoAnimEffect.msoAnimEffectAscend:
                case PowerPoint.MsoAnimEffect.msoAnimEffectFloat:
                    effectName = "float_up";
                    category = "entrance";
                    break;
                case PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink:
                    effectName = "grow_shrink";
                    category = "emphasis";
                    break;
                case PowerPoint.MsoAnimEffect.msoAnimEffectSpin:
                    effectName = "spin";
                    category = "emphasis";
                    break;
                default:
                    effectName = "appear";
                    category = isExit ? "exit" : "entrance";
                    break;
            }

            if (!isExit
                && (type == PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink
                    || type == PowerPoint.MsoAnimEffect.msoAnimEffectSpin))
            {
                category = "emphasis";
            }
        }

        private static bool TryParseDir(string dir, out PowerPoint.MsoAnimDirection direction, out string error)
        {
            direction = PowerPoint.MsoAnimDirection.msoAnimDirectionLeft;
            error = null;
            string d = (dir ?? "").Trim().ToLowerInvariant();
            switch (d)
            {
                case "left":
                    direction = PowerPoint.MsoAnimDirection.msoAnimDirectionLeft;
                    return true;
                case "right":
                    direction = PowerPoint.MsoAnimDirection.msoAnimDirectionRight;
                    return true;
                case "top":
                case "up":
                    direction = PowerPoint.MsoAnimDirection.msoAnimDirectionUp;
                    return true;
                case "bottom":
                case "down":
                    direction = PowerPoint.MsoAnimDirection.msoAnimDirectionDown;
                    return true;
                default:
                    error = "dir 须为 left|right|top|bottom";
                    return false;
            }
        }

        private static PowerPoint.MsoAnimDirection? TryReadDirection(PowerPoint.Effect effect)
        {
            try
            {
                return effect.EffectParameters.Direction;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string DirToName(PowerPoint.MsoAnimDirection? direction)
        {
            if (!direction.HasValue)
            {
                return null;
            }

            switch (direction.Value)
            {
                case PowerPoint.MsoAnimDirection.msoAnimDirectionLeft:
                    return "left";
                case PowerPoint.MsoAnimDirection.msoAnimDirectionRight:
                    return "right";
                case PowerPoint.MsoAnimDirection.msoAnimDirectionUp:
                case PowerPoint.MsoAnimDirection.msoAnimDirectionTop:
                    return "top";
                case PowerPoint.MsoAnimDirection.msoAnimDirectionDown:
                case PowerPoint.MsoAnimDirection.msoAnimDirectionBottom:
                    return "bottom";
                default:
                    return null;
            }
        }
    }
}
