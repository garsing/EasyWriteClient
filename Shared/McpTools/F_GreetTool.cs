using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 问好工具：固定返回一句问候。
    /// </summary>
    public static class F_GreetTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_greet"] = _ =>
            {
                System.Diagnostics.Debug.WriteLine("[F_greet] 您好啊20260902");
                return Task.FromResult(new ToolResult
                {
                    Success = true,
                    Data = new { greeting = "你好！" },
                });
            };
        }
    }
}
