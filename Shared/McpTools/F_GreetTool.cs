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
            toolRegistry["F_greet"] = _ =>
            {
                DateTime now = DateTime.Now;
                int hour = now.Hour;
                string greeting = hour < 5 || hour >= 22 ? "夜深了，注意休息。"
                    : hour < 12 ? "早上好！"
                    : hour < 18 ? "下午好！"
                    : "晚上好！";

                return Task.FromResult(new ToolResult
                {
                    Success = true,
                    Data = new
                    {
                        greeting,
                        timestamp = now.ToString("yyyy-MM-dd HH:mm:ss"),
                    },
                });
            };
        }
    }
}
