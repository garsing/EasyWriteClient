using WordAddIn1.OpenFiles;

namespace WordAddIn1
{
    public static class EtApplicationResolver
    {
        public static bool TryResolve(
            out object application,
            out string error,
            bool createIfMissing = false,
            bool makeVisible = true)
        {
            application = EtCom.TryGetActiveApplication(out _);
            if (application != null)
            {
                if (makeVisible)
                {
                    EtCom.EnsureVisible(application);
                }

                error = null;
                return true;
            }

            if (!createIfMissing)
            {
                error = "WPS 表格不可用；请先启动 WPS 表格。";
                return false;
            }

            return EtCom.TryCreateApplication(out application, out error);
        }
    }
}
