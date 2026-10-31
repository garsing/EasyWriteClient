using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordAddIn1.BrowserHost;

namespace BrowserNavigateTest
{
    /// <summary>本机网页：点出 confirm / prompt / alert，工具必须在数秒内返回。</summary>
    internal static class ScriptDialogLive
    {
        public static int Run()
        {
            string html = Path.GetFullPath(Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "..", "..", "..", "..",
                "Test", "BrowserNavigateTest", "fixtures", "script-dialog", "index.html"));
            if (!File.Exists(html))
            {
                html = Path.GetFullPath(Path.Combine(
                    FindRepo(),
                    "EasyWriteClient", "Test", "BrowserNavigateTest", "fixtures", "script-dialog", "index.html"));
            }

            if (!File.Exists(html))
            {
                Console.WriteLine("FAIL 找不到 " + html);
                return 1;
            }

            Application.EnableVisualStyles();
            int code = 1;
            var finished = new ManualResetEvent(false);
            ThreadPool.QueueUserWorkItem(_ =>
            {
                if (!finished.WaitOne(25000))
                {
                    Console.WriteLine("FAIL 25 秒仍未结束，点击卡在确认框里");
                    Environment.Exit(2);
                }
            });

            var form = new YiWriteBrowserForm();
            form.StartPosition = FormStartPosition.Manual;
            form.Location = new System.Drawing.Point(40, 40);
            form.Shown += async (s, e) =>
            {
                try
                {
                    await RunAsync(form, new Uri(html).AbsoluteUri).ConfigureAwait(true);
                    code = 0;
                    Console.WriteLine("PASS script-dialog-live");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("FAIL " + ex.Message);
                    code = 1;
                }
                finally
                {
                    finished.Set();
                    form.BeginInvoke(new Action(form.Close));
                }
            };

            Application.Run(form);
            return code;
        }

        private static async Task RunAsync(YiWriteBrowserForm form, string url)
        {
            string tab = Guid.NewGuid().ToString("N");
            Console.WriteLine("open " + url);
            await form.NavigateAsync(url, true, tab).ConfigureAwait(true);
            form.PinToolPage(tab);
            await Task.Delay(400).ConfigureAwait(true);

            await ClickNamedAsync(form, "撤回").ConfigureAwait(true);
            YiWriteBrowserPage page = form.TryGetPage(tab);
            BrowserScriptDialogSession session = page != null ? page.PendingScriptDialog : null;
            if (session == null || !session.IsPending)
            {
                throw new InvalidOperationException("点撤回后没有挂起的确认框");
            }

            string tree = BrowserScriptDialog.Render(session).TreeText;
            Console.WriteLine(tree.Replace("\r\n", "\n"));
            if (tree.IndexOf("确定撤回该工时吗？", StringComparison.Ordinal) < 0
                || tree.IndexOf("确定", StringComparison.Ordinal) < 0
                || tree.IndexOf("取消", StringComparison.Ordinal) < 0)
            {
                throw new InvalidOperationException("确认框快照不对");
            }

            if (tree.IndexOf("ready", StringComparison.Ordinal) >= 0)
            {
                throw new InvalidOperationException("快照里还带着页面");
            }

            BrowserScriptDialog.Commit(session, BrowserScriptDialogStep.Accept);
            page.ClearScriptDialogIf(session);
            await WaitOutAsync(page, "accepted").ConfigureAwait(true);

            await ClickNamedAsync(form, "撤回").ConfigureAwait(true);
            session = page.PendingScriptDialog;
            if (session == null || !session.IsPending)
            {
                throw new InvalidOperationException("第二次撤回没有确认框");
            }

            BrowserScriptDialog.Commit(session, BrowserScriptDialogStep.Dismiss);
            page.ClearScriptDialogIf(session);
            await WaitOutAsync(page, "dismissed").ConfigureAwait(true);

            await ClickNamedAsync(form, "提问").ConfigureAwait(true);
            session = page.PendingScriptDialog;
            if (session == null || session.Kind != BrowserScriptDialogKind.Prompt)
            {
                throw new InvalidOperationException("提问没有 prompt");
            }

            session.SetText("乙");
            BrowserScriptDialog.Commit(session, BrowserScriptDialogStep.Accept);
            page.ClearScriptDialogIf(session);
            await WaitOutAsync(page, "prompt:乙").ConfigureAwait(true);

            await ClickNamedAsync(form, "提示").ConfigureAwait(true);
            session = page.PendingScriptDialog;
            if (session == null || session.Kind != BrowserScriptDialogKind.Alert)
            {
                throw new InvalidOperationException("提示没有 alert");
            }

            BrowserScriptDialog.Commit(session, BrowserScriptDialogStep.Accept);
            page.ClearScriptDialogIf(session);
            await WaitOutAsync(page, "alerted").ConfigureAwait(true);
            form.UnpinToolPage();
        }

        private static async Task ClickNamedAsync(YiWriteBrowserForm form, string name)
        {
            int nodeId = 0;
            for (int i = 0; i < 10 && nodeId == 0; i++)
            {
                string json = await form.GetAccessibilityTreeJsonAsync().ConfigureAwait(true);
                BrowserAxBuildResult built = BrowserAxTreeBuilder.BuildOverview(json);
                if (built.Success)
                {
                    foreach (BrowserRefEntry entry in built.Refs.Values)
                    {
                        if (entry.BackendDomNodeId.HasValue
                            && string.Equals(entry.Name, name, StringComparison.Ordinal)
                            && string.Equals(entry.Role, "button", StringComparison.OrdinalIgnoreCase))
                        {
                            nodeId = entry.BackendDomNodeId.Value;
                            break;
                        }
                    }
                }

                if (nodeId == 0)
                {
                    await Task.Delay(200).ConfigureAwait(true);
                }
            }

            if (nodeId == 0)
            {
                throw new InvalidOperationException("找不到按钮 " + name);
            }

            Console.WriteLine("click " + name + " node=" + nodeId);
            Task click = BrowserInteractEngine.ClickAsync(form, nodeId, null, "pointer", "button");
            Task done = await Task.WhenAny(click, Task.Delay(8000)).ConfigureAwait(true);
            if (!ReferenceEquals(done, click))
            {
                throw new InvalidOperationException("点击「" + name + "」8 秒没有返回");
            }

            await click.ConfigureAwait(true);
            Console.WriteLine("click " + name + " returned");
        }

        private static async Task WaitOutAsync(YiWriteBrowserPage page, string expect)
        {
            for (int i = 0; i < 20; i++)
            {
                string raw = await page.Core.ExecuteScriptAsync(
                    "document.getElementById('out').textContent").ConfigureAwait(true);
                string text = (raw ?? "").Trim().Trim('"');
                Console.WriteLine("out=" + text);
                if (string.Equals(text, expect, StringComparison.Ordinal))
                {
                    return;
                }

                await Task.Delay(150).ConfigureAwait(true);
            }

            throw new InvalidOperationException("页面结果不是 " + expect);
        }

        private static string FindRepo()
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            for (int i = 0; i < 8 && !string.IsNullOrEmpty(dir); i++)
            {
                if (Directory.Exists(Path.Combine(dir, "EasyWriteClient")))
                {
                    return dir;
                }

                dir = Path.GetDirectoryName(dir);
            }

            return AppDomain.CurrentDomain.BaseDirectory;
        }
    }
}
