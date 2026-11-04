using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>问好：人名、中/英、当前时段。</summary>
    public static class F_GreetTool
    {
        static readonly string[] ZhPeriods = { "早晨", "中午", "下午", "晚上", "深夜" };
        static readonly string[] EnPeriods = { "morning", "noon", "afternoon", "evening", "night" };
        static readonly string[] ZhHellos = { "你好", "早上好", "中午好", "下午好", "晚上好", "深夜好" };
        static readonly string[] EnHellos = { "Hello", "Good morning", "Good afternoon", "Good afternoon", "Good evening", "Good night" };

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_greet"] = args =>
            {
                try
                {
                    var name = Arg(args, "name");
                    if (string.IsNullOrWhiteSpace(name)) name = "朋友";
                    var en = string.Equals(Arg(args, "language"), "en", StringComparison.OrdinalIgnoreCase);
                    var timed = Flag(args, "use_time_greeting", true);
                    var now = DateTime.Now;
                    var period = Period(now.Hour);
                    var hello = (en ? EnHellos : ZhHellos)[timed ? period + 1 : 0];
                    return Task.FromResult(new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            greeting = en ? hello + ", " + name + "!" : hello + "，" + name + "！",
                            name,
                            language = en ? "en" : "zh",
                            use_time_greeting = timed,
                            time_of_day = (en ? EnPeriods : ZhPeriods)[period],
                            date = now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                            time = now.ToString("HH:mm:ss", CultureInfo.InvariantCulture),
                        },
                    });
                }
                catch (Exception ex)
                {
                    return Task.FromResult(new ToolResult { Success = false, Error = "问候失败: " + ex.Message });
                }
            };
        }

        static int Period(int hour)
        {
            if (hour >= 5 && hour < 12) return 0;
            if (hour >= 12 && hour < 14) return 1;
            if (hour >= 14 && hour < 18) return 2;
            if (hour >= 18 && hour < 22) return 3;
            return 4;
        }

        static string Arg(Dictionary<string, object> args, string key)
        {
            object v;
            return args != null && args.TryGetValue(key, out v) && v != null ? v.ToString().Trim() : null;
        }

        static bool Flag(Dictionary<string, object> args, string key, bool fallback)
        {
            object v;
            if (args == null || !args.TryGetValue(key, out v) || v == null) return fallback;
            if (v is bool b) return b;
            bool parsed;
            return bool.TryParse(v.ToString(), out parsed) ? parsed : fallback;
        }
    }
}
