using System;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// ClickMode 选树与操作后端；track 只说明页面在哪。
    /// </summary>
    internal sealed class BrowserOpMode
    {
        public string Engine { get; private set; }

        public string ClickMode { get; private set; }

        public bool IsAttach { get; private set; }

        public bool IsPointer
        {
            get
            {
                return string.Equals(ClickMode, "pointer", StringComparison.Ordinal);
            }
        }

        public static BrowserOpMode Resolve(BrowserChannel channel)
        {
            bool attach = channel != null
                && string.Equals(channel.Track, "attach", StringComparison.OrdinalIgnoreCase);
            string engine = attach && channel != null
                ? ExtensionHost.GetAttachBrowserKind(channel.TabUuid)
                : "yiwrite";
            return new BrowserOpMode
            {
                Engine = ConfigManager.NormalizeBrowserEngine(engine),
                ClickMode = ConfigManager.GetBrowserClickMode(engine),
                IsAttach = attach
            };
        }
    }
}
