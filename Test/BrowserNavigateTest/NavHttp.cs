using System;
using System.Net;
using System.Text;
using System.Threading;

namespace BrowserNavigateTest
{
    /// <summary>本机 http 夹具。navigate 只收 http/https，不用 file://。</summary>
    internal sealed class NavHttp : IDisposable
    {
        private readonly HttpListener _listener;
        private readonly Thread _thread;
        private volatile bool _stop;

        public NavHttp()
        {
            int port = 18791;
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

            One = "http://127.0.0.1:" + Port.ToString() + "/one";
            Two = "http://127.0.0.1:" + Port.ToString() + "/two";
            _thread = new Thread(Listen) { IsBackground = true, Name = "nav-http" };
            _thread.Start();
        }

        public int Port { get; private set; }

        public string One { get; }

        public string Two { get; }

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

                string path = (ctx.Request.Url?.AbsolutePath ?? "/").TrimEnd('/').ToLowerInvariant();
                string title = path.EndsWith("two", StringComparison.Ordinal) ? "EasyWriteNavTwo" : "EasyWriteNavOne";
                string body = "<!doctype html><html><head><meta charset=\"utf-8\"><title>"
                    + title
                    + "</title></head><body><p>"
                    + title
                    + "</p></body></html>";
                byte[] bytes = Encoding.UTF8.GetBytes(body);
                ctx.Response.StatusCode = 200;
                ctx.Response.ContentType = "text/html; charset=utf-8";
                ctx.Response.ContentLength64 = bytes.Length;
                try
                {
                    ctx.Response.OutputStream.Write(bytes, 0, bytes.Length);
                }
                catch
                {
                }

                try
                {
                    ctx.Response.Close();
                }
                catch
                {
                }
            }
        }
    }
}
