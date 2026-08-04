using System;
using System.Windows.Forms;

namespace EasyWriteClient.Desktop
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            // B2：Desktop 宿主禁用 ActiveDocument 回退（须先 open/create 得渠道）
            WordAddIn1.ChannelHost.Kind = WordAddIn1.ChannelHostKind.Desktop;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
