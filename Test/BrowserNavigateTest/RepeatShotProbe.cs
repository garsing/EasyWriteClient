using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordAddIn1.BrowserHost;

namespace BrowserNavigateTest
{
    /// <summary>自建页走出重复段标题，再对这个 ref 做局部截图。</summary>
    internal static class RepeatShotProbe
    {
        public static int Run()
        {
            string htmlPath = Path.GetFullPath(Path.Combine(
                FindRepo(),
                "EasyWriteClient", "Test", "BrowserNavigateTest", "fixtures", "repeat-shot", "index.html"));
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine("FAIL 找不到 " + htmlPath);
                return 1;
            }

            string html = File.ReadAllText(htmlPath, Encoding.UTF8);
            Application.EnableVisualStyles();
            string data = Path.Combine(Path.GetTempPath(), "EasyWriteRepeatShot", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(data);
            YiWriteBrowserForm.UserDataFolderOverride = data;

            int code = 1;
            var finished = new ManualResetEvent(false);
            ThreadPool.QueueUserWorkItem(_ =>
            {
                if (!finished.WaitOne(45000))
                {
                    Console.WriteLine("FAIL 45 秒仍未结束");
                    Environment.Exit(2);
                }
            });

            using (var site = new OnePageSite(html))
            {
                var pump = new Form
                {
                    ShowInTaskbar = false,
                    WindowState = FormWindowState.Minimized
                };
                pump.Shown += async (s, e) =>
                {
                    try
                    {
                        code = await RunAsync(site.Url).ConfigureAwait(true);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("FAIL " + ex);
                        code = 1;
                    }
                    finally
                    {
                        finished.Set();
                        pump.BeginInvoke(new Action(() => Application.Exit()));
                    }
                };
                Application.Run(pump);
            }

            try
            {
                Directory.Delete(data, true);
            }
            catch
            {
            }

            return code;
        }

        private static async Task<int> RunAsync(string url)
        {
            Console.WriteLine("open " + url);
            BrowserNavigateResult nav = await YiWriteBrowserHost.NavigateAsync(url, true, null).ConfigureAwait(true);
            if (nav == null || !nav.Success || nav.Channel == null)
            {
                Console.WriteLine("FAIL 开页 " + (nav == null ? "" : nav.Error));
                return 1;
            }

            await Task.Delay(800).ConfigureAwait(true);
            BrowserSnapshotResult snap = await YiWriteBrowserHost.SnapshotAsync(nav.Channel, null).ConfigureAwait(true);
            if (snap == null || !snap.Success)
            {
                Console.WriteLine("FAIL 快照 " + (snap == null ? "" : snap.Error));
                return 1;
            }

            string tree = snap.Snapshot ?? "";
            Console.WriteLine(tree.Replace("\r\n", "\n"));
            string repeatRef = FindRepeatRef(tree);
            string headingRef = FindNamedRef(tree, "课题来源");
            if (string.IsNullOrEmpty(repeatRef) || string.IsNullOrEmpty(headingRef))
            {
                Console.WriteLine("FAIL 缺 repeat 或标题 ref");
                return 1;
            }

            Console.WriteLine("capture heading ref=" + headingRef + " 不传 up");
            BrowserCaptureResult own = await BrowserHostAdapter
                .CaptureLocalAsync(nav.Channel, headingRef, null)
                .ConfigureAwait(true);
            bool ownOk = own != null
                && own.Success
                && own.Image != null
                && own.Image.Height > 0
                && own.Image.Height < 400;
            Console.WriteLine(
                "heading success=" + (own != null && own.Success)
                + " error=" + (own == null ? "" : own.Error)
                + " size=" + (own != null && own.Image != null ? own.Image.Width + "x" + own.Image.Height : ""));
            if (!ownOk)
            {
                Console.WriteLine("FAIL 标题没有裁出自己的盒子");
                return 1;
            }

            Console.WriteLine("capture repeat ref=" + repeatRef + " 不传 up");
            BrowserCaptureResult repeatOwn = await BrowserHostAdapter
                .CaptureLocalAsync(nav.Channel, repeatRef, null)
                .ConfigureAwait(true);
            string repeatError = repeatOwn == null ? "" : (repeatOwn.Error ?? "");
            Console.WriteLine("repeat-own success=" + (repeatOwn != null && repeatOwn.Success) + " error=" + repeatError);
            if (repeatOwn == null || repeatOwn.Success || repeatError.IndexOf("没有可截区域", StringComparison.Ordinal) < 0)
            {
                Console.WriteLine("FAIL 重复段标题应没有自己的盒子");
                return 1;
            }

            Console.WriteLine("capture repeat ref=" + repeatRef + " up=1");
            BrowserCaptureResult parent = await BrowserHostAdapter
                .CaptureLocalAsync(nav.Channel, repeatRef, 1)
                .ConfigureAwait(true);
            bool parentOk = parent != null
                && parent.Success
                && parent.Image != null
                && parent.Image.Height > own.Image.Height;
            Console.WriteLine(
                "repeat-up success=" + (parent != null && parent.Success)
                + " error=" + (parent == null ? "" : parent.Error)
                + " size=" + (parent != null && parent.Image != null ? parent.Image.Width + "x" + parent.Image.Height : ""));
            if (!parentOk)
            {
                Console.WriteLine("FAIL up=1 没有裁到比标题更大的上一行");
                return 1;
            }

            Console.WriteLine("PASS 默认裁本节点，up 再往上");
            return 0;
        }

        private static string FindNamedRef(string tree, string name)
        {
            string[] lines = (tree ?? "").Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in lines)
            {
                if (line.IndexOf(name, StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                string refId = RefOnLine(line);
                if (!string.IsNullOrEmpty(refId))
                {
                    return refId;
                }
            }

            return null;
        }

        private static string FindRepeatRef(string tree)
        {
            string[] lines = (tree ?? "").Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in lines)
            {
                if (line.IndexOf("repeat", StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                string refId = RefOnLine(line);
                if (!string.IsNullOrEmpty(refId))
                {
                    return refId;
                }
            }

            return null;
        }

        private static string RefOnLine(string line)
        {
            if (string.IsNullOrEmpty(line))
            {
                return null;
            }

            int at = line.IndexOf("ref=", StringComparison.Ordinal);
            if (at < 0)
            {
                return null;
            }

            int start = at + 4;
            int end = start;
            while (end < line.Length && (char.IsLetterOrDigit(line[end]) || line[end] == '_'))
            {
                end++;
            }

            if (end <= start)
            {
                return null;
            }

            return line.Substring(start, end - start);
        }

        private static string FindRepo()
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            for (int i = 0; i < 8 && !string.IsNullOrEmpty(dir); i++)
            {
                if (Directory.Exists(Path.Combine(dir, "EasyWriteClient"))
                    && Directory.Exists(Path.Combine(dir, "tests")))
                {
                    return dir;
                }

                dir = Path.GetDirectoryName(dir);
            }

            return Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", ".."));
        }

        private sealed class OnePageSite : IDisposable
        {
            private readonly HttpListener _listener;
            private readonly Thread _thread;
            private volatile bool _stop;

            public OnePageSite(string html)
            {
                int port = 18940;
                _listener = new HttpListener();
                for (int i = 0; i < 20; i++)
                {
                    _listener.Prefixes.Clear();
                    _listener.Prefixes.Add("http://127.0.0.1:" + port.ToString() + "/");
                    try
                    {
                        _listener.Start();
                        Url = "http://127.0.0.1:" + port.ToString() + "/index.html";
                        break;
                    }
                    catch (HttpListenerException)
                    {
                        port++;
                        if (i == 19)
                        {
                            throw;
                        }
                    }
                }

                _thread = new Thread(() => Listen(html)) { IsBackground = true, Name = "repeat-shot-http" };
                _thread.Start();
            }

            public string Url { get; private set; }

            public void Dispose()
            {
                _stop = true;
                try
                {
                    _listener.Stop();
                }
                catch
                {
                }

                try
                {
                    _listener.Close();
                }
                catch
                {
                }
            }

            private void Listen(string html)
            {
                byte[] bytes = Encoding.UTF8.GetBytes(html);
                while (!_stop)
                {
                    HttpListenerContext ctx;
                    try
                    {
                        ctx = _listener.GetContext();
                    }
                    catch
                    {
                        break;
                    }

                    try
                    {
                        ctx.Response.StatusCode = 200;
                        ctx.Response.ContentType = "text/html; charset=utf-8";
                        ctx.Response.ContentLength64 = bytes.Length;
                        ctx.Response.OutputStream.Write(bytes, 0, bytes.Length);
                        ctx.Response.OutputStream.Close();
                    }
                    catch
                    {
                    }
                }
            }
        }
    }
}
