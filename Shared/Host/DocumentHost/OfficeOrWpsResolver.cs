using WordAddIn1.HostPlatform;

namespace WordAddIn1.DocumentHost
{
    internal static class OfficeOrWpsResolver
    {
        public static bool TryResolve(string fullPath, string app, out OfficeOrWps vendor, out string error)
        {
            vendor = OfficeOrWps.Office;
            error = null;
            string mapped = app;
            if (string.IsNullOrWhiteSpace(mapped))
            {
                if (!AssociatedExecutable.TryMapDefaultApp(fullPath, out mapped, out error))
                {
                    return false;
                }
            }

            mapped = mapped.Trim().ToLowerInvariant();
            if (mapped == "word")
            {
                vendor = OfficeOrWps.Office;
                return true;
            }

            if (mapped == "wps")
            {
                vendor = OfficeOrWps.Wps;
                return true;
            }

            error = "文字文件不能用 app=" + mapped + " 打开";
            return false;
        }
    }
}
