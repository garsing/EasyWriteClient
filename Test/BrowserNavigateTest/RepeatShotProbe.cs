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
            if (string.IsNullOrEmpty(repeatRef))
            {
                Console.WriteLine("FAIL 这页没有折出 repeat");
                return 1;
            }

            Console.WriteLine("capture ref=" + repeatRef + " up=1");
            BrowserCaptureResult shot = await BrowserHostAdapter
                .CaptureLocalAsync(nav.Channel, repeatRef, 1)
                .ConfigureAwait(true);
            string error = shot == null ? "" : (shot.Error ?? "");
            Console.WriteLine("capture success=" + (shot != null && shot.Success) + " error=" + error);
            bool imageOk = shot != null
                && shot.Success
                && shot.Image != null
                && shot.Image.Bytes != null
                && shot.Image.Bytes.Length > 0
                && shot.Image.Width > 0
                && shot.Image.Height > 0;
            if (imageOk)
            {
                Console.WriteLine(
                    "PASS 重复段标题按上一行裁出图 "
                    + shot.Image.Width + "x" + shot.Image.Height);
                return 0;
            }

            Console.WriteLine("FAIL 重复段标题没有按上一行裁出图");
            return 1;
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

                int at = line.IndexOf("[ref=", StringComparison.Ordinal);
                if (at < 0)
                {
                    at = line.IndexOf("ref=", StringComparison.Ordinal);
                }

                if (at < 0)
                {
                    continue;
                }

                int start = line.IndexOf('=', at) + 1;
                int end = start;
                while (end < line.Length && (char.IsLetterOrDigit(line[end]) || line[end] == '_'))
                {
                    end++;
                }

                if (end > start)
                {
                    return line.Substring(start, end - start);
                }
            }

            return null;
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
