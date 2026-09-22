using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 问好工具：按人名、语言、语气和当前时段生成问候。
    /// </summary>
    public static class F_GreetTool
    {
        private static readonly string[] Languages = { "zh", "en", "ja", "ko", "fr", "de", "es" };
        private static readonly string[] Styles = { "casual", "formal", "friendly", "professional" };

        // 顺序：普通、早晨、中午、下午、晚上、深夜
        private static readonly Dictionary<string, string[]> Hellos =
            new Dictionary<string, string[]>(StringComparer.Ordinal)
            {
                ["zh"] = new[] { "你好", "早上好", "中午好", "下午好", "晚上好", "深夜好" },
                ["en"] = new[] { "Hello", "Good morning", "Good afternoon", "Good afternoon", "Good evening", "Good night" },
                ["ja"] = new[] { "こんにちは", "おはようございます", "こんにちは", "こんにちは", "こんばんは", "こんばんは" },
                ["ko"] = new[] { "안녕하세요", "좋은 아침입니다", "안녕하세요", "안녕하세요", "안녕하세요", "안녕하세요" },
                ["fr"] = new[] { "Bonjour", "Bonjour", "Bonjour", "Bonjour", "Bonsoir", "Bonne nuit" },
                ["de"] = new[] { "Guten Tag", "Guten Morgen", "Guten Tag", "Guten Tag", "Guten Abend", "Gute Nacht" },
                ["es"] = new[] { "Hola", "Buenos días", "Buenas tardes", "Buenas tardes", "Buenas tardes", "Buenas noches" },
            };

        private static readonly Dictionary<string, string[]> TimeLabels =
            new Dictionary<string, string[]>(StringComparer.Ordinal)
            {
                ["zh"] = new[] { "早晨", "中午", "下午", "晚上", "深夜" },
                ["en"] = new[] { "morning", "noon", "afternoon", "evening", "night" },
                ["ja"] = new[] { "朝", "昼", "午後", "夕方", "夜" },
                ["ko"] = new[] { "아침", "정오", "오후", "저녁", "밤" },
                ["fr"] = new[] { "matin", "midi", "après-midi", "soir", "nuit" },
                ["de"] = new[] { "Morgen", "Mittag", "Nachmittag", "Abend", "Nacht" },
                ["es"] = new[] { "mañana", "mediodía", "tarde", "noche", "noche" },
            };

        private static readonly Dictionary<string, string[]> SeasonLabels =
            new Dictionary<string, string[]>(StringComparer.Ordinal)
            {
                ["zh"] = new[] { "春季", "夏季", "秋季", "冬季" },
                ["en"] = new[] { "spring", "summer", "autumn", "winter" },
                ["ja"] = new[] { "春", "夏", "秋", "冬" },
                ["ko"] = new[] { "봄", "여름", "가을", "겨울" },
                ["fr"] = new[] { "printemps", "été", "automne", "hiver" },
                ["de"] = new[] { "Frühling", "Sommer", "Herbst", "Winter" },
                ["es"] = new[] { "primavera", "verano", "otoño", "invierno" },
            };

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

                    string language = Normalize(GetString(args, "language"), Languages, "zh");
                    string style = Normalize(GetString(args, "style"), Styles, "friendly");
                    bool useTimeGreeting = GetBool(args, "use_time_greeting", true);

                    DateTime now = DateTime.Now;
                    int period = PeriodIndex(now.Hour);
                    string greeting = BuildGreeting(name, language, style, useTimeGreeting, period);

                    return Task.FromResult(new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            greeting,
                            name,
                            language,
                            style,
                            use_time_greeting = useTimeGreeting,
                            time_of_day = TimeLabels[language][period],
                            season = SeasonLabels[language][SeasonIndex(now.Month)],
                            day_of_week = now.ToString("dddd", CultureOf(language)),
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

        private static string BuildGreeting(string name, string language, string style, bool timed, int period)
        {
            string hello = Hellos[language][timed ? period + 1 : 0];
            string who = language == "ja" ? name + "さん" : language == "ko" ? name + "님" : name;
            bool polite = style == "formal" || style == "professional";

            if (language == "zh")
            {
                if (style == "casual" && !timed) return "嗨，" + who + "！";
                if (style == "professional") return hello + "，" + who + "，很高兴为您服务。";
                if (style == "friendly") return hello + "，" + who + "！很高兴见到你！";
                return hello + "，" + who + (polite ? "。" : "！");
            }

            string sep = language == "ja" ? "、" : ", ";
            string end = polite ? "." : "!";
            if (language == "ja" || language == "ko")
            {
                end = polite ? "。" : "！";
            }

            if (style == "casual" && !timed && language == "en")
            {
                return "Hey, " + who + "!";
            }

            return hello + sep + who + end;
        }

        private static int PeriodIndex(int hour)
        {
            if (hour >= 5 && hour < 12) return 0;
            if (hour >= 12 && hour < 14) return 1;
            if (hour >= 14 && hour < 18) return 2;
            if (hour >= 18 && hour < 22) return 3;
            return 4;
        }

        private static int SeasonIndex(int month)
        {
            if (month >= 3 && month <= 5) return 0;
            if (month >= 6 && month <= 8) return 1;
            if (month >= 9 && month <= 11) return 2;
            return 3;
        }

        private static CultureInfo CultureOf(string language)
        {
            switch (language)
            {
                case "en": return CultureInfo.GetCultureInfo("en-US");
                case "ja": return CultureInfo.GetCultureInfo("ja-JP");
                case "ko": return CultureInfo.GetCultureInfo("ko-KR");
                case "fr": return CultureInfo.GetCultureInfo("fr-FR");
                case "de": return CultureInfo.GetCultureInfo("de-DE");
                case "es": return CultureInfo.GetCultureInfo("es-ES");
                default: return CultureInfo.GetCultureInfo("zh-CN");
            }
        }

        private static string Normalize(string value, string[] allowed, string fallback)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return fallback;
            }

            string normalized = value.Trim().ToLowerInvariant();
            return allowed.Contains(normalized) ? normalized : fallback;
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
