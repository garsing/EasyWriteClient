using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading;
using System.Web;

namespace BrowserNavigateTest
{
    /// <summary>本机 http 挂 tests/ 与 真实html。navigate 只用 http。</summary>
    internal sealed class FixtureHttp : IDisposable
    {
        private readonly HttpListener _listener;
        private readonly Thread _thread;
        private readonly Dictionary<string, string> _roots =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private volatile bool _stop;

        public FixtureHttp()
        {
            string repo = FindRepoRoot();
            _roots["iframe"] = Path.Combine(repo, "tests", "iframe");
            _roots["center-tabs"] = Path.Combine(repo, "tests", "center-tabs");
            _roots["hover-card"] = Path.Combine(repo, "tests", "hover-card");
            _roots["native-select"] = Path.Combine(repo, "tests", "native-select");
            _roots["overlay-cookie"] = Path.Combine(repo, "tests", "overlay-cookie");
            _roots["file-chooser"] = Path.Combine(repo, "tests", "file-chooser");
            _roots["disabled-input"] = Path.Combine(repo, "tests", "disabled-input");
            _roots["overflow-scroll"] = Path.Combine(repo, "tests", "overflow-scroll");
            _roots["d18-risk"] = Path.Combine(repo, "tests", "d18-risk");
            _roots["form-controls"] = Path.Combine(repo, "tests", "form-controls");
            _roots["form-extra"] = Path.Combine(repo, "tests", "form-extra");
            _roots["public-example"] = Path.Combine(repo, "tests", "public-example");
            _roots["real"] = Path.Combine(
                repo, "docs", "测试", "浏览器测试", "snapshot与iframe", "真实html");

            int port = 18792;
            _listener = new HttpListener();
            for (int i = 0; i < 20; i++)
            {
                _listener.Prefixes.Clear();
                _listener.Prefixes.Add("http://127.0.0.1:" + port.ToString() + "/");
                try
                {
                    _listener.Start();
                    Port = port;
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

            _thread = new Thread(Listen) { IsBackground = true, Name = "snp-http" };
            _thread.Start();
        }

        public int Port { get; private set; }

        public string Url(string mount, string file)
        {
            return "http://127.0.0.1:" + Port.ToString() + "/" + mount.Trim('/') + "/" + file.TrimStart('/');
        }

        public string Real(string alias)
        {
            string dir = _roots["real"];
            string name = ResolveRealName(dir, alias);
            return Url("real", name);
        }

        public void Dispose()
        {
            _stop = true;
            try { _listener.Stop(); } catch { }
            try { _listener.Close(); } catch { }
        }

        private static string ResolveRealName(string dir, string alias)
        {
            if (alias == "train")
            {
                return "项目实战去实训.html";
            }

            if (alias == "square")
            {
                return "项目广场.html";
            }

            if (alias == "type")
            {
                return "项目类别.html";
            }

            if (alias == "jd-home")
            {
                return FindStartsWith(dir, "京东(JD");
            }

            if (alias == "jd-mine")
            {
                return FindStartsWith(dir, "京东商城");
            }

            return alias;
        }

        private static string FindStartsWith(string dir, string prefix)
        {
            if (!Directory.Exists(dir))
            {
                return prefix + ".html";
            }

            foreach (string f in Directory.GetFiles(dir, "*.html"))
            {
                if (Path.GetFileName(f).StartsWith(prefix, StringComparison.Ordinal))
                {
                    return Path.GetFileName(f);
                }
            }

            return prefix + ".html";
        }

        private static string FindRepoRoot()
        {
            string d = AppDomain.CurrentDomain.BaseDirectory ?? "";
            for (int i = 0; i < 10; i++)
            {
                if (Directory.Exists(Path.Combine(d, "tests", "iframe"))
                    && Directory.Exists(Path.Combine(d, "docs", "测试", "浏览器测试")))
                {
                    return d;
                }

                string parent = Path.GetDirectoryName(d);
                if (string.IsNullOrEmpty(parent) || parent == d)
                {
                    break;
                }

                d = parent;
            }

            throw new InvalidOperationException("找不到仓库根（tests/iframe）");
        }

        private void Listen()
        {
            while (!_stop && _listener.IsListening)
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
                    Serve(ctx);
                }
                catch
                {
                    try { ctx.Response.Abort(); } catch { }
                }
            }
        }

        private void Serve(HttpListenerContext ctx)
        {
            string raw = ctx.Request.Url == null ? "/" : ctx.Request.Url.AbsolutePath;
            string path = HttpUtility.UrlDecode(raw).Trim('/').Replace('/', Path.DirectorySeparatorChar);
            string[] parts = path.Split(new[] { Path.DirectorySeparatorChar }, 2);
            if (parts.Length < 2 || !_roots.TryGetValue(parts[0], out string root))
            {
                ctx.Response.StatusCode = 404;
                ctx.Response.Close();
                return;
            }

            string full = Path.GetFullPath(Path.Combine(root, parts[1]));
            string rootFull = Path.GetFullPath(root);
            if (!full.StartsWith(rootFull, StringComparison.OrdinalIgnoreCase) || !File.Exists(full))
            {
                ctx.Response.StatusCode = 404;
                ctx.Response.Close();
                return;
            }

            byte[] bytes = File.ReadAllBytes(full);
            ctx.Response.StatusCode = 200;
            ctx.Response.ContentType = "text/html; charset=utf-8";
            ctx.Response.ContentLength64 = bytes.Length;
            ctx.Response.OutputStream.Write(bytes, 0, bytes.Length);
            ctx.Response.Close();
        }
    }
}
