using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 问好工具：按人名、语言、语气和当前时段生成问候，并返回时间上下文。
    /// </summary>
    public static class F_GreetTool
    {
        private static readonly string[] SupportedLanguages = { "zh", "en", "ja", "ko", "fr", "de", "es" };
        private static readonly string[] SupportedStyles = { "casual", "formal", "friendly", "professional" };

        private static readonly Dictionary<string, Dictionary<string, string>> TimeOfDayTranslations =
            new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal)
            {
                ["morning"] = LangMap("早晨", "morning", "朝", "아침", "matin", "Morgen", "mañana"),
                ["noon"] = LangMap("中午", "noon", "昼", "정오", "midi", "Mittag", "mediodía"),
                ["afternoon"] = LangMap("下午", "afternoon", "午後", "오후", "après-midi", "Nachmittag", "tarde"),
                ["evening"] = LangMap("晚上", "evening", "夕方", "저녁", "soir", "Abend", "noche"),
                ["night"] = LangMap("深夜", "night", "夜", "밤", "nuit", "Nacht", "noche"),
            };

        private static readonly Dictionary<string, Dictionary<string, string>> SeasonTranslations =
            new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal)
            {
                ["spring"] = LangMap("春季", "spring", "春", "봄", "printemps", "Frühling", "primavera"),
                ["summer"] = LangMap("夏季", "summer", "夏", "여름", "été", "Sommer", "verano"),
                ["autumn"] = LangMap("秋季", "autumn", "秋", "가을", "automne", "Herbst", "otoño"),
                ["winter"] = LangMap("冬季", "winter", "冬", "겨울", "hiver", "Winter", "invierno"),
            };

        private static readonly Dictionary<DayOfWeek, Dictionary<string, string>> DayOfWeekTranslations =
            new Dictionary<DayOfWeek, Dictionary<string, string>>
            {
                [DayOfWeek.Monday] = LangMap("星期一", "Monday", "月曜日", "월요일", "Lundi", "Montag", "Lunes"),
                [DayOfWeek.Tuesday] = LangMap("星期二", "Tuesday", "火曜日", "화요일", "Mardi", "Dienstag", "Martes"),
                [DayOfWeek.Wednesday] = LangMap("星期三", "Wednesday", "水曜日", "수요일", "Mercredi", "Mittwoch", "Miércoles"),
                [DayOfWeek.Thursday] = LangMap("星期四", "Thursday", "木曜日", "목요일", "Jeudi", "Donnerstag", "Jueves"),
                [DayOfWeek.Friday] = LangMap("星期五", "Friday", "金曜日", "금요일", "Vendredi", "Freitag", "Viernes"),
                [DayOfWeek.Saturday] = LangMap("星期六", "Saturday", "土曜日", "토요일", "Samedi", "Samstag", "Sábado"),
                [DayOfWeek.Sunday] = LangMap("星期日", "Sunday", "日曜日", "일요일", "Dimanche", "Sonntag", "Domingo"),
            };

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_greet"] = args =>
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("[F_greet] 您好啊哈哈20260905");

                    string name = GetStringArg(args, "name");
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        name = "朋友";
                    }

                    string language = NormalizeChoice(GetStringArg(args, "language"), SupportedLanguages, "zh");
                    string style = NormalizeChoice(GetStringArg(args, "style"), SupportedStyles, "friendly");
                    bool useTimeGreeting = GetBoolArg(args, "use_time_greeting", true);

                    DateTime now = DateTime.Now;
                    string periodKey = GetTimePeriodKey(now.Hour);
                    string timeOfDay = Translate(TimeOfDayTranslations, periodKey, language);
                    string season = Translate(SeasonTranslations, GetSeasonKey(now.Month), language);
                    string greeting = GenerateGreeting(name, language, style, useTimeGreeting, periodKey);

                    System.Diagnostics.Debug.WriteLine(
                        "[F_greet] " + greeting + " hour=" + now.Hour + " lang=" + language + " style=" + style);

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
                            time_of_day = timeOfDay,
                            season,
                            day_of_week = TranslateDay(now.DayOfWeek, language),
                            timestamp = now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
                            date = now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                            time = now.ToString("HH:mm:ss", CultureInfo.InvariantCulture),
                            hour = now.Hour,
                            month = now.Month,
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

        private static string GenerateGreeting(
            string name,
            string language,
            string style,
            bool useTimeGreeting,
            string periodKey)
        {
            if (!useTimeGreeting)
            {
                return GetPlainGreeting(name, language, style);
            }

            return GetTimedGreeting(name, language, style, periodKey);
        }

        private static string GetPlainGreeting(string name, string language, string style)
        {
            switch (language)
            {
                case "en":
                    return Pick(style, "Hey, " + name + "!", "Hello, " + name + "!", "Hello, " + name + "! Nice to meet you!", "Hello, " + name + ". Nice to meet you.");
                case "ja":
                    return Pick(style, "やあ、" + name + "！", "こんにちは、" + name + "さん！", "こんにちは、" + name + "さん！", "はじめまして、" + name + "さん。");
                case "ko":
                    return Pick(style, "안녕, " + name + "!", "안녕하세요, " + name + "님!", "안녕하세요, " + name + "님!", "안녕하세요, " + name + "님.");
                case "fr":
                    return Pick(style, "Salut, " + name + " !", "Bonjour, " + name + " !", "Bonjour, " + name + " !", "Bonjour, " + name + ".");
                case "de":
                    return Pick(style, "Hallo, " + name + "!", "Guten Tag, " + name + "!", "Guten Tag, " + name + "!", "Guten Tag, " + name + ".");
                case "es":
                    return Pick(style, "¡Hola, " + name + "!", "¡Hola, " + name + "!", "¡Hola, " + name + "!", "Buenos días, " + name + ".");
                default:
                    return Pick(style, "嗨，" + name + "！", "您好，" + name + "！", "你好，" + name + "！很高兴见到你！", "您好，" + name + "，很高兴为您服务。");
            }
        }

        private static string GetTimedGreeting(string name, string language, string style, string periodKey)
        {
            switch (language)
            {
                case "en":
                    return GetTimedGreetingEn(name, style, periodKey);
                case "ja":
                    return GetTimedGreetingJa(name, style, periodKey);
                case "ko":
                    return GetTimedGreetingKo(name, style, periodKey);
                case "fr":
                    return GetTimedGreetingFr(name, style, periodKey);
                case "de":
                    return GetTimedGreetingDe(name, style, periodKey);
                case "es":
                    return GetTimedGreetingEs(name, style, periodKey);
                default:
                    return GetTimedGreetingZh(name, style, periodKey);
            }
        }

        private static string GetTimedGreetingZh(string name, string style, string periodKey)
        {
            switch (periodKey)
            {
                case "morning":
                    return Pick(style, "早啊，" + name + "！", "早上好，" + name + "！", "早上好，" + name + "！新的一天开始了！", "您好，" + name + "，新的一天开始了。");
                case "noon":
                    return Pick(style, "中午好，" + name + "！", "中午好，" + name + "！", "中午好，" + name + "！用餐愉快！", "您好，" + name + "，午间愉快。");
                case "afternoon":
                    return Pick(style, "下午好，" + name + "！", "下午好，" + name + "！", "下午好，" + name + "！工作顺利！", "您好，" + name + "，下午工作顺利。");
                case "evening":
                    return Pick(style, "晚上好，" + name + "！", "晚上好，" + name + "！", "晚上好，" + name + "！今天过得怎么样？", "您好，" + name + "，晚间愉快。");
                default:
                    return Pick(style, "这么晚了，" + name + "，还在忙呢？", "深夜好，" + name + "，请注意休息。", "晚上好，" + name + "！这么晚了还在工作，辛苦了！", "您好，" + name + "，深夜时分请注意休息。");
            }
        }

        private static string GetTimedGreetingEn(string name, string style, string periodKey)
        {
            switch (periodKey)
            {
                case "morning":
                    return Pick(style, "Morning, " + name + "!", "Good morning, " + name + "!", "Good morning, " + name + "! Have a wonderful day!", "Good morning, " + name + ". Have a great day.");
                case "noon":
                    return Pick(style, "Hey " + name + "!", "Good afternoon, " + name + "!", "Good afternoon, " + name + "! Hope you're having a great day!", "Good afternoon, " + name + ". Enjoy your lunch.");
                case "afternoon":
                    return Pick(style, "Hey " + name + "!", "Good afternoon, " + name + "!", "Good afternoon, " + name + "! How's your day going?", "Good afternoon, " + name + ". How can I assist you?");
                case "evening":
                    return Pick(style, "Evening, " + name + "!", "Good evening, " + name + "!", "Good evening, " + name + "! How was your day?", "Good evening, " + name + ". How may I help you?");
                default:
                    return Pick(style, "Late night, " + name + "! Still up?", "Good night, " + name + ". Please rest well.", "Good evening, " + name + "! It's late, hope you're doing well!", "Good evening, " + name + ". It's late, please take care.");
            }
        }

        private static string GetTimedGreetingJa(string name, string style, string periodKey)
        {
            switch (periodKey)
            {
                case "morning":
                    return Pick(style, "おはよう、" + name + "！", "おはようございます、" + name + "さん！", "おはようございます、" + name + "さん！良い一日を！", "おはようございます、" + name + "さん。よろしくお願いします。");
                case "evening":
                case "night":
                    return Pick(style, "こんばんは、" + name + "！", "こんばんは、" + name + "さん！", "こんばんは、" + name + "さん！今日はどうでしたか？", "こんばんは、" + name + "さん。お疲れ様です。");
                default:
                    return Pick(style, "こんにちは、" + name + "！", "こんにちは、" + name + "さん！", "こんにちは、" + name + "さん！お会いできて嬉しいです！", "こんにちは、" + name + "さん。お疲れ様です。");
            }
        }

        private static string GetTimedGreetingKo(string name, string style, string periodKey)
        {
            switch (periodKey)
            {
                case "morning":
                    return Pick(style, "좋은 아침, " + name + "!", "좋은 아침입니다, " + name + "님!", "좋은 아침입니다, " + name + "님! 좋은 하루 보내세요!", "좋은 아침입니다, " + name + "님. 좋은 하루 되세요.");
                case "evening":
                case "night":
                    return Pick(style, "안녕, " + name + "! 저녁 잘 보내고 있어?", "안녕하세요, " + name + "님! 좋은 저녁 되세요.", "안녕하세요, " + name + "님! 오늘 하루는 어땠어요?", "안녕하세요, " + name + "님. 오늘 하루 수고하셨습니다.");
                default:
                    return Pick(style, "안녕, " + name + "!", "안녕하세요, " + name + "님!", "안녕하세요, " + name + "님! 만나서 반갑습니다!", "안녕하세요, " + name + "님. 반갑습니다.");
            }
        }

        private static string GetTimedGreetingFr(string name, string style, string periodKey)
        {
            switch (periodKey)
            {
                case "morning":
                    return Pick(style, "Salut, " + name + " !", "Bonjour, " + name + " !", "Bonjour, " + name + " ! Belle journée !", "Bonjour, " + name + ". Comment puis-je vous aider ?");
                case "evening":
                    return Pick(style, "Bonsoir, " + name + " !", "Bonsoir, " + name + " !", "Bonsoir, " + name + " ! Comment s'est passée la journée ?", "Bonsoir, " + name + ". Comment puis-je vous aider ?");
                case "night":
                    return Pick(style, "Il est tard, " + name + " !", "Bonne nuit, " + name + ".", "Bonne nuit, " + name + ", reposez-vous bien !", "Bonne nuit, " + name + ". Prenez soin de vous.");
                default:
                    return Pick(style, "Salut, " + name + " !", "Bonjour, " + name + " !", "Bonjour, " + name + " ! Ravi de vous rencontrer !", "Bonjour, " + name + ". Comment puis-je vous aider ?");
            }
        }

        private static string GetTimedGreetingDe(string name, string style, string periodKey)
        {
            switch (periodKey)
            {
                case "morning":
                    return Pick(style, "Morgen, " + name + "!", "Guten Morgen, " + name + "!", "Guten Morgen, " + name + "! Schönen Tag!", "Guten Morgen, " + name + ". Wie kann ich Ihnen helfen?");
                case "evening":
                    return Pick(style, "Abend, " + name + "!", "Guten Abend, " + name + "!", "Guten Abend, " + name + "! Wie war dein Tag?", "Guten Abend, " + name + ". Wie kann ich Ihnen helfen?");
                case "night":
                    return Pick(style, "Schon spät, " + name + "!", "Gute Nacht, " + name + ".", "Gute Nacht, " + name + ", ruh dich gut aus!", "Gute Nacht, " + name + ". Bitte ruhen Sie sich aus.");
                default:
                    return Pick(style, "Hallo, " + name + "!", "Guten Tag, " + name + "!", "Guten Tag, " + name + "! Freut mich, Sie kennenzulernen!", "Guten Tag, " + name + ". Wie kann ich Ihnen helfen?");
            }
        }

        private static string GetTimedGreetingEs(string name, string style, string periodKey)
        {
            switch (periodKey)
            {
                case "morning":
                    return Pick(style, "¡Hola, " + name + "!", "¡Buenos días, " + name + "!", "¡Buenos días, " + name + "! ¡Que tengas un buen día!", "Buenos días, " + name + ". ¿En qué puedo ayudarle?");
                case "evening":
                    return Pick(style, "¡Buenas, " + name + "!", "¡Buenas tardes, " + name + "!", "¡Buenas tardes, " + name + "! ¿Cómo te fue el día?", "Buenas tardes, " + name + ". ¿En qué puedo ayudarle?");
                case "night":
                    return Pick(style, "Es tarde, " + name + ".", "Buenas noches, " + name + ".", "¡Buenas noches, " + name + "! Descansa bien.", "Buenas noches, " + name + ". Cuídese.");
                default:
                    return Pick(style, "¡Hola, " + name + "!", "¡Buenas tardes, " + name + "!", "¡Hola, " + name + "! ¡Encantado de conocerte!", "Buenos días, " + name + ". ¿En qué puedo ayudarle?");
            }
        }

        private static string Pick(string style, string casual, string formal, string friendly, string professional)
        {
            switch (style)
            {
                case "casual":
                    return casual;
                case "formal":
                    return formal;
                case "professional":
                    return professional;
                default:
                    return friendly;
            }
        }

        private static string GetTimePeriodKey(int hour)
        {
            if (hour >= 5 && hour < 12) return "morning";
            if (hour >= 12 && hour < 14) return "noon";
            if (hour >= 14 && hour < 18) return "afternoon";
            if (hour >= 18 && hour < 22) return "evening";
            return "night";
        }

        private static string GetSeasonKey(int month)
        {
            if (month >= 3 && month <= 5) return "spring";
            if (month >= 6 && month <= 8) return "summer";
            if (month >= 9 && month <= 11) return "autumn";
            return "winter";
        }

        private static string Translate(
            Dictionary<string, Dictionary<string, string>> table,
            string key,
            string language)
        {
            Dictionary<string, string> row;
            if (!table.TryGetValue(key, out row))
            {
                return key;
            }

            string text;
            if (row.TryGetValue(language, out text) && !string.IsNullOrEmpty(text))
            {
                return text;
            }

            return row["zh"];
        }

        private static string TranslateDay(DayOfWeek day, string language)
        {
            Dictionary<string, string> row;
            if (!DayOfWeekTranslations.TryGetValue(day, out row))
            {
                return day.ToString();
            }

            string text;
            if (row.TryGetValue(language, out text) && !string.IsNullOrEmpty(text))
            {
                return text;
            }

            return row["zh"];
        }

        private static Dictionary<string, string> LangMap(
            string zh, string en, string ja, string ko, string fr, string de, string es)
        {
            return new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["zh"] = zh,
                ["en"] = en,
                ["ja"] = ja,
                ["ko"] = ko,
                ["fr"] = fr,
                ["de"] = de,
                ["es"] = es,
            };
        }

        private static string NormalizeChoice(string value, string[] allowed, string fallback)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return fallback;
            }

            string normalized = value.Trim().ToLowerInvariant();
            return allowed.Contains(normalized) ? normalized : fallback;
        }

        private static string GetStringArg(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            return args[key].ToString().Trim();
        }

        private static bool GetBoolArg(Dictionary<string, object> args, string key, bool fallback)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return fallback;
            }

            object raw = args[key];
            if (raw is bool)
            {
                return (bool)raw;
            }

            bool parsed;
            if (bool.TryParse(raw.ToString(), out parsed))
            {
                return parsed;
            }

            return fallback;
        }
    }
}
