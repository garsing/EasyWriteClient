using System;
using System.IO;
using Microsoft.Win32;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 幂等写入 Native Messaging Host 登记（Chrome + Edge）。扩展 ID 来自配置或侧载说明。
    /// </summary>
    public static class NativeMessagingRegistration
    {
        public const string HostName = "com.yiwrite.browser_bridge";

        public static void EnsureRegistered()
        {
            string bridgeExe = ResolveBridgeExePath();
            if (string.IsNullOrWhiteSpace(bridgeExe) || !File.Exists(bridgeExe))
            {
                System.Diagnostics.Debug.WriteLine(
                    "[NM] YiWriteBrowserBridge.exe 未找到: " + bridgeExe);
                return;
            }

            string manifestPath = WriteHostManifest(bridgeExe);
            RegisterForBrowser(
                @"Software\Google\Chrome\NativeMessagingHosts\" + HostName,
                manifestPath);
            RegisterForBrowser(
                @"Software\Microsoft\Edge\NativeMessagingHosts\" + HostName,
                manifestPath);
        }

        private static string ResolveBridgeExePath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory ?? "";
            string nextToDesktop = Path.Combine(baseDir, "YiWriteBrowserBridge.exe");
            if (File.Exists(nextToDesktop))
            {
                return nextToDesktop;
            }

            // 开发：从 Desktop 输出目录旁的 BrowserBridge
            string sibling = Path.GetFullPath(Path.Combine(
                baseDir, "..", "..", "..", "..", "BrowserBridge", "bin", "Debug", "net472", "YiWriteBrowserBridge.exe"));
            if (File.Exists(sibling))
            {
                return sibling;
            }

            sibling = Path.GetFullPath(Path.Combine(
                baseDir, "..", "..", "..", "..", "BrowserBridge", "bin", "Release", "net472", "YiWriteBrowserBridge.exe"));
            if (File.Exists(sibling))
            {
                return sibling;
            }

            return nextToDesktop;
        }

        private static string WriteHostManifest(string bridgeExe)
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "YiWrite",
                "native-messaging");
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, HostName + ".json");

            string[] allowed = LoadAllowedExtensionOrigins();
            var origins = new System.Text.StringBuilder();
            origins.Append("[");
            for (int i = 0; i < allowed.Length; i++)
            {
                if (i > 0)
                {
                    origins.Append(",");
                }

                origins.Append("\"").Append(EscapeJson(allowed[i])).Append("\"");
            }

            origins.Append("]");

            string json =
                "{\n" +
                "  \"name\": \"" + HostName + "\",\n" +
                "  \"description\": \"YiWrite browser bridge\",\n" +
                "  \"path\": \"" + EscapeJson(bridgeExe.Replace('\\', '/')) + "\",\n" +
                "  \"type\": \"stdio\",\n" +
                "  \"allowed_origins\": " + origins + "\n" +
                "}\n";

            // Windows NM path 须用反斜杠完整路径
            json =
                "{\r\n" +
                "  \"name\": \"" + HostName + "\",\r\n" +
                "  \"description\": \"YiWrite browser bridge\",\r\n" +
                "  \"path\": \"" + EscapeJson(bridgeExe) + "\",\r\n" +
                "  \"type\": \"stdio\",\r\n" +
                "  \"allowed_origins\": " + origins + "\r\n" +
                "}\r\n";

            File.WriteAllText(path, json);
            return path;
        }

        private static string[] LoadAllowedExtensionOrigins()
        {
            var list = new System.Collections.Generic.List<string>();
            string baseDir = AppDomain.CurrentDomain.BaseDirectory ?? "";
            string cfg = Path.Combine(baseDir, "browser-extension-id.txt");
            if (File.Exists(cfg))
            {
                foreach (string line in File.ReadAllLines(cfg))
                {
                    string id = (line ?? "").Trim();
                    if (id.Length == 0 || id.StartsWith("#", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (!id.StartsWith("chrome-extension://", StringComparison.OrdinalIgnoreCase))
                    {
                        id = "chrome-extension://" + id.TrimEnd('/') + "/";
                    }

                    list.Add(id);
                }
            }

            // 侧载时若尚未写入 ID，放宽占位（登记脚本会覆盖）
            if (list.Count == 0)
            {
                list.Add("chrome-extension://INVALID_REPLACE_ME/");
            }

            return list.ToArray();
        }

        private static void RegisterForBrowser(string relativeKey, string manifestPath)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(relativeKey))
                {
                    key?.SetValue("", manifestPath, RegistryValueKind.String);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[NM] registry " + relativeKey + ": " + ex.Message);
            }
        }

        private static string EscapeJson(string s)
        {
            if (s == null)
            {
                return "";
            }

            return s.Replace("\\", "\\\\").Replace("\"", "\\\"");
        }
    }
}
