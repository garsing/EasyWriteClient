using System;
using System.Windows.Forms;
using WordAddIn1;

namespace EasyWriteClient.Desktop
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            // B2：Desktop 宿主禁用 ActiveDocument 回退（须先 open/create 得渠道）
            ChannelHost.Kind = ChannelHostKind.Desktop;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 启用 open_files 等 DebugCategories 时，镜像到 Desktop/logs/
            EasyWriteLog.Initialize("Desktop");
            try
            {
                Application.Run(new MainForm());
            }
            finally
            {
                EasyWriteLog.Shutdown();
            }
        }
    }
}
