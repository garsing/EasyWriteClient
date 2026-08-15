using System;
using System.Reflection;
using WordAddIn1.OpenFiles;

namespace WordAddIn1
{
    public static class WpsApplicationResolver
    {
        public static bool TryResolve(
            out object application,
            out string error,
            bool createIfMissing = false)
        {
            application = WpsCom.TryGetActiveApplication(out _);
            if (application != null)
            {
                error = null;
                TrySetVisible(application);
                return true;
            }

            if (!createIfMissing)
            {
                error = "WPS 文字不可用；请先启动 WPS 文字。";
                return false;
            }

            foreach (string id in WpsCom.ProgIds)
            {
                try
                {
                    Type t = Type.GetTypeFromProgID(id);
                    if (t == null)
                    {
                        continue;
                    }

                    application = Activator.CreateInstance(t);
                    if (application == null)
                    {
                        continue;
                    }

                    TrySetVisible(application);
                    error = null;
                    return true;
                }
                catch (Exception)
                {
                    application = null;
                }
            }

            error = "无法创建 WPS 文字（Kwps.Application / wps.Application），是否未安装 WPS？";
            return false;
        }

        private static void TrySetVisible(object app)
        {
            try
            {
                app.GetType().InvokeMember(
                    "Visible",
                    BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    app,
                    new object[] { true });
            }
            catch (Exception)
            {
            }
        }
    }
}
