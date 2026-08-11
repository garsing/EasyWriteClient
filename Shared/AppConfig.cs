using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 应用程序统一配置文件类
    /// </summary>
    public class AppConfig
    {
        public ApiConfig Api { get; set; }
        public LLMApiConfig LLMApi { get; set; }
        public AppSettings App { get; set; }
        public UiSettings UI { get; set; }
        public DocumentColorsSettings DocumentColors { get; set; }
        public DocumentNavigateSettings DocumentNavigate { get; set; }
        public DocumentActionsAmbiguitySettings DocumentActionsAmbiguity { get; set; }
        public WorkspaceSettings Workspace { get; set; }
        /// <summary>侧栏「打开文件」可监视应用白名单（无总开关）。</summary>
        public OpenFilesSettings OpenFiles { get; set; }

        public AppConfig()
        {
            // 不在这里设置默认值，默认值将在配置文件中定义
        }

        public class ApiConfig
        {
            public string BaseUrl { get; set; }
            public int TimeoutSeconds { get; set; }
            public int MaxRetries { get; set; }
            public Dictionary<string, string> Endpoints { get; set; }

            public ApiConfig()
            {
                Endpoints = new Dictionary<string, string>();
            }
        }

        public class LLMApiConfig
        {
            public string BaseUrl { get; set; }
            public int TimeoutSeconds { get; set; }
            public Dictionary<string, string> Endpoints { get; set; }

            public LLMApiConfig()
            {
                Endpoints = new Dictionary<string, string>();
            }
        }

        public class AppSettings
        {
            public string Name { get; set; }
            public string Version { get; set; }

            /// <summary>
            /// 分类 Debug 白名单；含 "all" 则全部输出；空 / 未配置则关闭。
            /// </summary>
            public List<string> DebugCategories { get; set; }

            /// <summary>
            /// readText 字符数超过此阈值时走后端 API（curr_doc_chunk_and_snapshot），否则本地 ProcessDocumentLocally。默认 60000。
            /// </summary>
            public int BackendProcessTextLengthThreshold { get; set; } = 60000;

            /// <summary>
            /// 插件启动时将 Word ActivePrinter 设为 Microsoft Print to PDF（或 XPS），避免 PageSetup 连离线打印机挂死。默认 true。
            /// </summary>
            public bool PreferPdfActivePrinterOnStartup { get; set; } = true;

            public AppSettings()
            {
                DebugCategories = new List<string>();
            }
        }

        public class UiSettings
        {
            public string Theme { get; set; }
            public string Language { get; set; }
        }

        public class DocumentColorsSettings
        {
            public string StrikethroughBackground { get; set; }
            public string InsertBackground { get; set; }
            public string AcceptButtonBackground { get; set; }
            public string RejectButtonBackground { get; set; }
        }

        public class DocumentNavigateSettings
        {
            public bool AutoNavigateAfterModify { get; set; } = true;
            public string NavigateAnchor { get; set; } = "first_new_sentence";
        }

        public class DocumentActionsAmbiguitySettings
        {
            public int ContextLeftSentences { get; set; } = 2;
            public int ContextRightSentences { get; set; } = 2;
        }

        public class WorkspaceSettings
        {
            public string Mode { get; set; } = "sessions";
        }

        /// <summary>打开文件探测配置；仅 Apps 白名单，无 Enabled 总开关。</summary>
        public class OpenFilesSettings
        {
            public List<OpenFilesAppEntry> Apps { get; set; }

            public OpenFilesSettings()
            {
                Apps = new List<OpenFilesAppEntry>();
            }
        }

        public class OpenFilesAppEntry
        {
            /// <summary>应用类型键：如 "word"、"wps"（小写）。</summary>
            public string Type { get; set; }
            public string DisplayName { get; set; }
            public List<string> ProcessNames { get; set; }
            public List<string> Extensions { get; set; }

            public OpenFilesAppEntry()
            {
                ProcessNames = new List<string>();
                Extensions = new List<string>();
            }
        }
    }

    /// <summary>
    /// 配置管理器
    /// </summary>
    public static class ConfigManager
    {
        private static readonly string ConfigFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
        private static AppConfig _config;
        private static readonly object _lock = new object();

        /// <summary>
        /// 获取配置实例
        /// </summary>
        public static AppConfig Config
        {
            get
            {
                if (_config == null)
                {
                    lock (_lock)
                    {
                        if (_config == null)
                        {
                            LoadConfig();
                        }
                    }
                }
                return _config;
            }
        }

        /// <summary>
        /// 加载配置文件
        /// </summary>
        private static void LoadConfig()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    string json = File.ReadAllText(ConfigFilePath, Encoding.UTF8);

                    try
                    {
                        _config = JsonConvert.DeserializeObject<AppConfig>(json);
                        NormalizeLoadedConfig(_config);
                    }
                    catch (Exception)
                    {
                        _config = CreateDefaultConfig();
                    }
                }
                else
                {
                    _config = CreateDefaultConfig();
                    SaveConfig(); // 创建默认配置文件
                }
            }
            catch (Exception)
            {
                _config = CreateDefaultConfig();
            }

            EasyWriteDiagnostics.RefreshFromConfig();
        }

        private static void NormalizeLoadedConfig(AppConfig config)
        {
            if (config == null)
            {
                return;
            }

            if (config.App == null)
            {
                config.App = new AppConfig.AppSettings();
            }

            if (config.App.DebugCategories == null)
            {
                config.App.DebugCategories = new List<string>();
            }

            // I12：OpenFiles / Apps 缺失或空 = 谁都不探测（不自动填白名单）
            if (config.OpenFiles == null)
            {
                config.OpenFiles = new AppConfig.OpenFilesSettings
                {
                    Apps = new List<AppConfig.OpenFilesAppEntry>()
                };
            }
            else if (config.OpenFiles.Apps == null)
            {
                config.OpenFiles.Apps = new List<AppConfig.OpenFilesAppEntry>();
            }
            else
            {
                foreach (var app in config.OpenFiles.Apps)
                {
                    if (app == null)
                    {
                        continue;
                    }

                    if (app.ProcessNames == null)
                    {
                        app.ProcessNames = new List<string>();
                    }

                    if (app.Extensions == null)
                    {
                        app.Extensions = new List<string>();
                    }
                }
            }
        }

        /// <summary>首次生成 / 随包推荐配置：预置 word + wps（不是 Normalize 兜底）。</summary>
        internal static AppConfig.OpenFilesSettings CreateDefaultOpenFilesSettings()
        {
            return new AppConfig.OpenFilesSettings
            {
                Apps = new List<AppConfig.OpenFilesAppEntry>
                {
                    new AppConfig.OpenFilesAppEntry
                    {
                        Type = "word",
                        DisplayName = "Word",
                        ProcessNames = new List<string> { "WINWORD" },
                        Extensions = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm" }
                    },
                    new AppConfig.OpenFilesAppEntry
                    {
                        Type = "wps",
                        DisplayName = "WPS文字",
                        ProcessNames = new List<string> { "wps" },
                        Extensions = new List<string> { ".doc", ".docx", ".wps", ".wpt" }
                    }
                }
            };
        }

        /// <summary>
        /// 创建默认配置
        /// </summary>
        private static AppConfig CreateDefaultConfig()
        {
            return new AppConfig
            {
                Api = new AppConfig.ApiConfig
                {
                    // 与仓库 config.json 对齐；ClickOnce 若未打进 config.json 时的兜底
                    BaseUrl = "https://addin-api.yi-write.com",
                    TimeoutSeconds = 30,
                    MaxRetries = 3,
                    Endpoints = new Dictionary<string, string>
                    {
                        { "Login", "/auth/login" },
                        { "LoginSms", "/auth/login/sms" },
                        { "Register", "/auth/register" },
                        { "SmsSend", "/auth/sms/send" },
                        { "Refresh", "/auth/refresh" },
                        { "Health", "/health" }
                    }
                },
                LLMApi = new AppConfig.LLMApiConfig
                {
                    BaseUrl = "https://addin-api.yi-write.com",
                    TimeoutSeconds = 30,
                    Endpoints = new Dictionary<string, string>
                    {
                        { "ChatCompletions", "/llm/chat/completions" }
                    }
                },
                App = new AppConfig.AppSettings
                {
                    Name = "WordAddIn1",
                    Version = "1.0.0",
                    DebugCategories = new List<string>(),
                    BackendProcessTextLengthThreshold = 60000,
                    PreferPdfActivePrinterOnStartup = true
                },
                UI = new AppConfig.UiSettings
                {
                    Theme = "Black",
                    Language = "zh-CN"
                },
                DocumentColors = new AppConfig.DocumentColorsSettings
                {
                    StrikethroughBackground = "wdColorGray05",
                    InsertBackground = "0xC0FFFF",
                    AcceptButtonBackground = "wdColorLightGreen",
                    RejectButtonBackground = "0xFF99FF"
                },
                DocumentNavigate = new AppConfig.DocumentNavigateSettings
                {
                    AutoNavigateAfterModify = true,
                    NavigateAnchor = "first_new_sentence"
                },
                DocumentActionsAmbiguity = new AppConfig.DocumentActionsAmbiguitySettings
                {
                    ContextLeftSentences = 2,
                    ContextRightSentences = 2
                },
                Workspace = new AppConfig.WorkspaceSettings
                {
                    Mode = "sessions"
                },
                OpenFiles = CreateDefaultOpenFilesSettings()
            };
        }

        public static AppConfig.DocumentActionsAmbiguitySettings GetDocumentActionsAmbiguitySettings()
        {
            var settings = Config.DocumentActionsAmbiguity;
            if (settings == null)
            {
                return new AppConfig.DocumentActionsAmbiguitySettings
                {
                    ContextLeftSentences = 2,
                    ContextRightSentences = 2
                };
            }

            return new AppConfig.DocumentActionsAmbiguitySettings
            {
                ContextLeftSentences = ClampAmbiguityContextCount(settings.ContextLeftSentences),
                ContextRightSentences = ClampAmbiguityContextCount(settings.ContextRightSentences)
            };
        }

        private static int ClampAmbiguityContextCount(int value)
        {
            if (value < 0)
            {
                return 0;
            }

            if (value > 5)
            {
                return 5;
            }

            return value;
        }

        /// <summary>
        /// 修改后定位配置（缺省节时等同默认值）
        /// </summary>
        public static AppConfig.DocumentNavigateSettings GetDocumentNavigateSettings()
        {
            var settings = Config.DocumentNavigate;
            if (settings == null)
            {
                return new AppConfig.DocumentNavigateSettings
                {
                    AutoNavigateAfterModify = true,
                    NavigateAnchor = "first_new_sentence"
                };
            }

            return settings;
        }

        /// <summary>
        /// 保存配置到文件
        /// </summary>
        public static bool SaveConfig()
        {
            try
            {
                string json = JsonConvert.SerializeObject(_config, Formatting.Indented);
                File.WriteAllText(ConfigFilePath, json, Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"保存配置文件失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 更新配置并保存
        /// </summary>
        public static bool UpdateConfig(AppConfig newConfig)
        {
            _config = newConfig;
            return SaveConfig();
        }

        /// <summary>
        /// 获取配置文件路径
        /// </summary>
        public static string GetConfigFilePath()
        {
            return ConfigFilePath;
        }

        /// <summary>
        /// 获取配置状态信息
        /// </summary>
        public static string GetConfigStatus()
        {
            var status = new StringBuilder();
            status.AppendLine("=== 应用程序配置状态 ===");
            status.AppendLine($"配置文件存在: {File.Exists(ConfigFilePath)}");
            status.AppendLine($"配置文件路径: {ConfigFilePath}");
            status.AppendLine($"API地址: {Config.Api.BaseUrl}");
            status.AppendLine($"超时时间: {Config.Api.TimeoutSeconds}秒");
            status.AppendLine($"最大重试次数: {Config.Api.MaxRetries}");
            status.AppendLine($"应用程序名称: {Config.App.Name}");
            status.AppendLine($"版本: {Config.App.Version}");
            var cats = Config.App.DebugCategories ?? new List<string>();
            status.AppendLine($"DebugCategories: [{string.Join(", ", cats)}]");
            status.AppendLine($"主题: {Config.UI.Theme}");
            status.AppendLine($"语言: {Config.UI.Language}");

            if (Config.DocumentColors != null)
            {
                status.AppendLine($"删除线背景色: {Config.DocumentColors.StrikethroughBackground}");
                status.AppendLine($"插入背景色: {Config.DocumentColors.InsertBackground}");
                status.AppendLine($"接受按钮背景色: {Config.DocumentColors.AcceptButtonBackground}");
                status.AppendLine($"丢弃按钮背景色: {Config.DocumentColors.RejectButtonBackground}");
            }
            else
            {
                status.AppendLine("文档颜色配置: 未设置");
            }

            return status.ToString();
        }

        /// <summary>
        /// 工作区读取模式（sessions / history），缺省 sessions。
        /// </summary>
        public static string GetWorkspaceMode()
        {
            var mode = Config?.Workspace?.Mode;
            if (string.IsNullOrWhiteSpace(mode))
            {
                return "sessions";
            }

            mode = mode.Trim().ToLowerInvariant();
            return mode == "history" ? "history" : "sessions";
        }

        /// <summary>
        /// 重载配置
        /// </summary>
        public static void ReloadConfig()
        {
            lock (_lock)
            {
                _config = null;
                LoadConfig();
            }
        }

        /// <summary>
        /// 测试配置系统
        /// </summary>
        public static string TestConfigSystem()
        {
            var testResults = new System.Text.StringBuilder();
            testResults.AppendLine("=== 配置系统测试 ===");

            try
            {
                // 测试配置加载
                var config = Config;
                testResults.AppendLine("✓ 配置加载成功");

                // 测试配置值
                testResults.AppendLine($"API地址: {config.Api.BaseUrl}");
                testResults.AppendLine($"超时时间: {config.Api.TimeoutSeconds}秒");
                testResults.AppendLine($"最大重试: {config.Api.MaxRetries}次");
                testResults.AppendLine($"应用名称: {config.App.Name}");
                testResults.AppendLine($"版本: {config.App.Version}");
                var cats = config.App.DebugCategories ?? new List<string>();
                testResults.AppendLine($"DebugCategories: [{string.Join(", ", cats)}]");
                testResults.AppendLine($"主题: {config.UI.Theme}");
                testResults.AppendLine($"语言: {config.UI.Language}");

                // 测试配置保存
                var originalUrl = config.Api.BaseUrl;
                config.Api.BaseUrl = "http://test.example.com";
                var saveResult = SaveConfig();
                testResults.AppendLine($"配置保存: {(saveResult ? "成功" : "失败")}");

                // 恢复原始配置
                config.Api.BaseUrl = originalUrl;
                SaveConfig();
                testResults.AppendLine("✓ 配置恢复成功");

                // 测试配置重载
                ReloadConfig();
                var reloadedConfig = Config;
                testResults.AppendLine($"配置重载: {(reloadedConfig.Api.BaseUrl == originalUrl ? "成功" : "失败")}");

                testResults.AppendLine("=== 测试完成 ===");
            }
            catch (Exception ex)
            {
                testResults.AppendLine($"✗ 测试失败: {ex.Message}");
            }

            return testResults.ToString();
        }

        /// <summary>
        /// 文档颜色配置帮助类
        /// </summary>
        public static class DocumentColorHelper
        {
            /// <summary>
            /// 将字符串转换为Word颜色
            /// </summary>
            public static Word.WdColor ParseColor(string colorString)
            {
                if (string.IsNullOrEmpty(colorString))
                {
                    System.Diagnostics.Debug.WriteLine($"[COLOR] 颜色字符串为空，返回默认颜色");
                    return Word.WdColor.wdColorAutomatic;
                }

                try
                {
                    // 处理预定义颜色
                    if (colorString.StartsWith("wdColor"))
                    {
                        var colorName = colorString;
                        switch (colorName)
                        {
                            case "wdColorGray05": return Word.WdColor.wdColorGray05;
                            case "wdColorGray10": return Word.WdColor.wdColorGray10;
                            case "wdColorGray25": return Word.WdColor.wdColorGray25;
                            case "wdColorGray50": return Word.WdColor.wdColorGray50;
                            case "wdColorLightGreen": return Word.WdColor.wdColorLightGreen;
                            case "wdColorYellow": return Word.WdColor.wdColorYellow;
                            case "wdColorLightYellow": return Word.WdColor.wdColorLightYellow;
                            case "wdColorAutomatic": return Word.WdColor.wdColorAutomatic;
                            default:
                                System.Diagnostics.Debug.WriteLine($"[COLOR] 未知的预定义颜色: {colorName}，返回默认颜色");
                                return Word.WdColor.wdColorAutomatic;
                        }
                    }
                    // 处理十六进制颜色
                    else if (colorString.StartsWith("0x"))
                    {
                        if (int.TryParse(colorString.Substring(2), System.Globalization.NumberStyles.HexNumber, null, out int colorValue))
                        {
                            return (Word.WdColor)colorValue;
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[COLOR] 无法解析十六进制颜色: {colorString}，返回默认颜色");
                            return Word.WdColor.wdColorAutomatic;
                        }
                    }
                    // 处理十进制数字
                    else if (int.TryParse(colorString, out int decimalValue))
                    {
                        return (Word.WdColor)decimalValue;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[COLOR] 无法识别的颜色格式: {colorString}，返回默认颜色");
                        return Word.WdColor.wdColorAutomatic;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[COLOR] 解析颜色失败: {colorString}, 错误: {ex.Message}，返回默认颜色");
                    return Word.WdColor.wdColorAutomatic;
                }
            }

            /// <summary>
            /// 获取删除线背景色
            /// </summary>
            public static Word.WdColor StrikethroughBackground => ParseColor(Config.DocumentColors?.StrikethroughBackground ?? "wdColorGray05");

            /// <summary>
            /// 获取插入内容背景色
            /// </summary>
            public static Word.WdColor InsertBackground => ParseColor(Config.DocumentColors?.InsertBackground ?? "0xC0FFFF");

            /// <summary>
            /// 获取接受按钮背景色
            /// </summary>
            public static Word.WdColor AcceptButtonBackground => ParseColor(Config.DocumentColors?.AcceptButtonBackground ?? "wdColorLightGreen");

            /// <summary>
            /// 获取丢弃按钮背景色
            /// </summary>
            public static Word.WdColor RejectButtonBackground => ParseColor(Config.DocumentColors?.RejectButtonBackground ?? "0xFF99FF");
        }
    }
}
