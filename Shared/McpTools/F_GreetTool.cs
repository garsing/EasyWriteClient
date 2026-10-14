using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 问好：按人名、中/英和当前时段生成一句问候。
    /// </summary>
    public static class F_GreetTool
    {
        // 早晨、中午、下午、晚上、深夜
        private static readonly string[] ZhPeriods = { "早晨", "中午", "下午", "晚上", "深夜" };
        private static readonly string[] EnPeriods = { "morning", "noon", "afternoon", "evening", "night" };

        // 下标 0 是不按时段的普通问候，其后与时段一一对应
        private static readonly string[] ZhHellos = { "你好", "早上好", "中午好", "下午好", "晚上好", "深夜好" };
        private static readonly string[] EnHellos = { "Hello", "Good morning", "Good afternoon", "Good afternoon", "Good evening", "Good night" };

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_greet"] = args =>
            {
                try
                {
                    string name = GetString(args, "name");
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        name = "朋友";
                    }

                    bool english = string.Equals(GetString(args, "language"), "en", StringComparison.OrdinalIgnoreCase);
                    bool timed = GetBool(args, "use_time_greeting", true);
                    DateTime now = DateTime.Now;
                    int period = PeriodIndex(now.Hour);

                    return Task.FromResult(new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            greeting = Greeting(name, english, timed, period),
                            name,
                            language = english ? "en" : "zh",
                            use_time_greeting = timed,
                            time_of_day = (english ? EnPeriods : ZhPeriods)[period],
                            date = now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                            time = now.ToString("HH:mm:ss", CultureInfo.InvariantCulture),
                        },
                    });
                }
                catch (Exception ex)
                {
                    return Task.FromResult(new ToolResult
                    {
                        Success = false,
                        Error = "问候失败: " + ex.Message,
                    });
                }
            };
        }

        private static string Greeting(string name, bool english, bool timed, int period)
        {
            string hello = (english ? EnHellos : ZhHellos)[timed ? period + 1 : 0];
            return english ? hello + ", " + name + "!" : hello + "，" + name + "！";
        }

        private static int PeriodIndex(int hour)
        {
            if (hour >= 5 && hour < 12) return 0;
            if (hour >= 12 && hour < 14) return 1;
            if (hour >= 14 && hour < 18) return 2;
            if (hour >= 18 && hour < 22) return 3;
            return 4;
        }

        private static string GetString(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            return args[key].ToString().Trim();
        }

        private static bool GetBool(Dictionary<string, object> args, string key, bool fallback)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return fallback;
            }

            if (args[key] is bool flag)
            {
                return flag;
            }

            bool parsed;
            return bool.TryParse(args[key].ToString(), out parsed) ? parsed : fallback;
        }
    }
}
