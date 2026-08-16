using WordAddIn1.HostPlatform;

namespace WordAddIn1.PresentationHost
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
            if (mapped == "powerpoint")
            {
                vendor = OfficeOrWps.Office;
                return true;
            }

            if (mapped == "wpp")
            {
                vendor = OfficeOrWps.Wps;
                return true;
            }

            error = "演示文稿不能用 app=" + mapped + " 打开（请用 powerpoint 或 wpp）";
            return false;
        }
    }
}
