using WordAddIn1.OpenFiles;

namespace WordAddIn1
{
    public static class WppApplicationResolver
    {
        public static bool TryResolve(
            out object application,
            out string error,
            bool createIfMissing = false,
            bool makeVisible = true)
        {
            application = WppCom.TryGetActiveApplication(out _);
            if (application != null)
            {
                if (makeVisible)
                {
                    WppCom.EnsureVisible(application);
                }

                error = null;
                return true;
            }

            if (!createIfMissing)
            {
                error = "WPS 演示不可用；请先启动 WPS 演示。";
                return false;
            }

            return WppCom.TryCreateApplication(out application, out error);
        }
    }
}
