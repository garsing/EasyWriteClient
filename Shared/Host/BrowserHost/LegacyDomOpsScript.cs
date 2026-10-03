using System;
using System.IO;
using System.Reflection;
using System.Text;

namespace WordAddIn1.BrowserHost
{
    internal static class LegacyDomOpsScript
    {
        private static string _cached;

        public static string Get()
        {
            if (!string.IsNullOrEmpty(_cached))
            {
                return _cached;
            }

            Assembly asm = typeof(LegacyDomOpsScript).Assembly;
            string name = "WordAddIn1.Host.BrowserHost.LegacyDomOps.js";
            Stream stream = asm.GetManifestResourceStream(name);
            if (stream == null)
            {
                foreach (string n in asm.GetManifestResourceNames())
                {
                    if (n != null && n.EndsWith("LegacyDomOps.js", StringComparison.OrdinalIgnoreCase))
                    {
                        stream = asm.GetManifestResourceStream(n);
                        break;
                    }
                }
            }

            if (stream == null)
            {
                throw new InvalidOperationException("未找到 LegacyDomOps.js");
            }

            using (stream)
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                _cached = reader.ReadToEnd();
            }

            return _cached;
        }
    }
}
