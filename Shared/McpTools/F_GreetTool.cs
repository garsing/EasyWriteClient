using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 问好工具：根据当前时段生成简单问候语。
    /// </summary>
    public static class F_GreetTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_greet"] = async (args) =>
            {
                try
                {
                    DateTime now = DateTime.Now;
                    string greeting = GetGreeting(now.Hour);

                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            greeting,
                            timestamp = now.ToString("yyyy-MM-dd HH:mm:ss"),
                        },
                    };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"问好失败: {ex.Message}" };
                }
            };
        }

        private static string GetGreeting(int hour)
        {
            if (hour >= 5 && hour < 12)
            {
                return "早上好！";
            }

            if (hour >= 12 && hour < 18)
            {
                EasyWriteDiagnostics.Log(DebugCategory.Tool, "[F_greet] 您好2026年7月25日啊");
                return "下午好！";
            }

            if (hour >= 18 && hour < 22)
            {
                return "晚上好！";
            }

            return "夜深了，注意休息。";
        }
    }
}
