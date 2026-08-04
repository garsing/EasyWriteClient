using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace WordAddIn1
{
    /// <summary>
    /// 独立隐藏 WebView2：SVG → PNG（≥2x）。须在 UI 线程调用。
    /// </summary>
    public static class SvgWebView2Rasterizer
    {
        public const double DefaultWidthCm = 14.0;
        public const int DefaultScale = 2;
        public const int DefaultTimeoutMs = 15000;

        private static readonly object Gate = new object();
        private static readonly Regex ViewBoxRegex = new Regex(
            @"viewBox\s*=\s*[""']\s*([0-9.+eE-]+)\s+([0-9.+eE-]+)\s+([0-9.+eE-]+)\s+([0-9.+eE-]+)\s*[""']",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public sealed class RasterizeResult
        {
            public bool Success { get; set; }
            public string Error { get; set; }
            public string PngPath { get; set; }
            public int WidthPx { get; set; }
            public int HeightPx { get; set; }
            public double WidthCm { get; set; }
            public double HeightCm { get; set; }
        }

        public static async Task<RasterizeResult> RasterizeAsync(
            string svgMarkup,
            double? widthCm = null,
            int scale = DefaultScale,
            int timeoutMs = DefaultTimeoutMs)
        {
            if (string.IsNullOrWhiteSpace(svgMarkup))
            {
                return Fail("render_failed: svg 为空");
            }

            double targetWidthCm = widthCm.HasValue && widthCm.Value > 0 ? widthCm.Value : DefaultWidthCm;
            if (scale < 1)
            {
                scale = DefaultScale;
            }

            ParseViewBox(svgMarkup, out double vbW, out double vbH);
            double aspect = vbH > 0 && vbW > 0 ? vbH / vbW : 1.0;
            double targetHeightCm = targetWidthCm * aspect;

            // 逻辑 CSS 像素（约 96 DPI：1cm ≈ 37.8px）
            int cssW = Math.Max(1, (int)Math.Round(targetWidthCm * 37.7952755906));
            int cssH = Math.Max(1, (int)Math.Round(targetHeightCm * 37.7952755906));
            int pixelW = cssW * scale;
            int pixelH = cssH * scale;

            string tempDir = Path.Combine(Path.GetTempPath(), "EasyWrite", "svg-raster");
            Directory.CreateDirectory(tempDir);
            string id = Guid.NewGuid().ToString("N");
            string htmlPath = Path.Combine(tempDir, $"raster-{id}.html");
            string pngPath = Path.Combine(tempDir, $"raster-{id}.png");

            try
            {
                File.WriteAllText(htmlPath, BuildHtml(svgMarkup, cssW, cssH), Encoding.UTF8);

                if (!Monitor.TryEnter(Gate, timeoutMs))
                {
                    return Fail("render_failed: 渲染器繁忙超时");
                }

                try
                {
                    using (var cts = new CancellationTokenSource(timeoutMs))
                    {
                        RasterizeResult result = await RasterizeWithWebViewAsync(
                            htmlPath,
                            pngPath,
                            cssW,
                            cssH,
                            pixelW,
                            pixelH,
                            targetWidthCm,
                            targetHeightCm,
                            cts.Token).ConfigureAwait(true);
                        return result;
                    }
                }
                finally
                {
                    Monitor.Exit(Gate);
                    TryDelete(htmlPath);
                }
            }
            catch (OperationCanceledException)
            {
                TryDelete(htmlPath);
                TryDelete(pngPath);
                return Fail("render_failed: 渲染超时");
            }
            catch (Exception ex)
            {
                TryDelete(htmlPath);
                TryDelete(pngPath);
                string hint = ex.Message != null && ex.Message.IndexOf("WebView2", StringComparison.OrdinalIgnoreCase) >= 0
                    ? "（请确认已安装 WebView2 Runtime）"
                    : string.Empty;
                return Fail($"render_failed: {ex.Message}{hint}");
            }
        }

        private static async Task<RasterizeResult> RasterizeWithWebViewAsync(
            string htmlPath,
            string pngPath,
            int cssW,
            int cssH,
            int pixelW,
            int pixelH,
            double widthCm,
            double heightCm,
            CancellationToken ct)
        {
            var form = new Form
            {
                ShowInTaskbar = false,
                FormBorderStyle = FormBorderStyle.None,
                Opacity = 0,
                Width = Math.Min(pixelW + 16, 2400),
                Height = Math.Min(pixelH + 16, 2400),
                StartPosition = FormStartPosition.Manual,
                Location = new Point(-16000, -16000)
            };

            var webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            form.Controls.Add(webView);
            form.Show();

            try
            {
                string userData = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordAddIn1",
                    "WebView2Data",
                    "SvgRaster");
                Directory.CreateDirectory(userData);

                var env = await CoreWebView2Environment.CreateAsync(userDataFolder: userData)
                    .ConfigureAwait(true);
                ct.ThrowIfCancellationRequested();
                await webView.EnsureCoreWebView2Async(env).ConfigureAwait(true);
                ct.ThrowIfCancellationRequested();

                webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                webView.CoreWebView2.Settings.AreDevToolsEnabled = false;

                var navTcs = new TaskCompletionSource<bool>();
                void OnNav(object sender, CoreWebView2NavigationCompletedEventArgs e)
                {
                    webView.CoreWebView2.NavigationCompleted -= OnNav;
                    if (e.IsSuccess)
                    {
                        navTcs.TrySetResult(true);
                    }
                    else
                    {
                        navTcs.TrySetException(new InvalidOperationException(
                            $"导航失败 WebErrorStatus={e.WebErrorStatus}"));
                    }
                }

                webView.CoreWebView2.NavigationCompleted += OnNav;
                string uri = new Uri(htmlPath).AbsoluteUri;
                webView.CoreWebView2.Navigate(uri);

                using (ct.Register(() => navTcs.TrySetCanceled()))
                {
                    await navTcs.Task.ConfigureAwait(true);
                }

                ct.ThrowIfCancellationRequested();

                // 稍等布局/字体
                await Task.Delay(120, ct).ConfigureAwait(true);

                string metricsJson =
                    $"{{\"width\":{cssW},\"height\":{cssH},\"deviceScaleFactor\":{(pixelW / (double)cssW).ToString(CultureInfo.InvariantCulture)},\"mobile\":false}}";
                await webView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                    "Emulation.setDeviceMetricsOverride",
                    metricsJson).ConfigureAwait(true);

                string shotJson = await webView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                    "Page.captureScreenshot",
                    "{\"format\":\"png\",\"fromSurface\":true}").ConfigureAwait(true);

                var serializer = new JavaScriptSerializer();
                var shotObj = serializer.Deserialize<System.Collections.Generic.Dictionary<string, object>>(shotJson);
                if (shotObj == null || !shotObj.ContainsKey("data") || shotObj["data"] == null)
                {
                    return Fail("render_failed: 截图无 data");
                }

                byte[] pngBytes = Convert.FromBase64String(shotObj["data"].ToString());
                File.WriteAllBytes(pngPath, pngBytes);

                System.Diagnostics.Debug.WriteLine(
                    $"[SvgRaster] ok css={cssW}x{cssH} px={pixelW}x{pixelH} cm={widthCm:F2}x{heightCm:F2} bytes={pngBytes.Length}");

                return new RasterizeResult
                {
                    Success = true,
                    PngPath = pngPath,
                    WidthPx = pixelW,
                    HeightPx = pixelH,
                    WidthCm = widthCm,
                    HeightCm = heightCm
                };
            }
            finally
            {
                try
                {
                    form.Controls.Remove(webView);
                    webView.Dispose();
                }
                catch
                {
                    // ignore
                }

                try
                {
                    form.Close();
                    form.Dispose();
                }
                catch
                {
                    // ignore
                }
            }
        }

        private static string BuildHtml(string svgMarkup, int cssW, int cssH)
        {
            // 不转义 SVG 本体（已是 XML）；仅包一层展示壳
            var sb = new StringBuilder();
            sb.Append("<!DOCTYPE html><html><head><meta charset=\"utf-8\"/>");
            sb.Append("<style>");
            sb.Append("html,body{margin:0;padding:0;background:#fff;");
            sb.Append("font-family:\"Microsoft YaHei\",\"Segoe UI\",sans-serif;}");
            sb.Append("svg{display:block;width:");
            sb.Append(cssW.ToString(CultureInfo.InvariantCulture));
            sb.Append("px;height:");
            sb.Append(cssH.ToString(CultureInfo.InvariantCulture));
            sb.Append("px;}");
            sb.Append("</style></head><body>");
            sb.Append(svgMarkup);
            sb.Append("</body></html>");
            return sb.ToString();
        }

        private static void ParseViewBox(string svg, out double w, out double h)
        {
            w = 680;
            h = 680;
            Match m = ViewBoxRegex.Match(svg ?? string.Empty);
            if (!m.Success)
            {
                return;
            }

            if (double.TryParse(m.Groups[3].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double vw)
                && double.TryParse(m.Groups[4].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double vh)
                && vw > 0 && vh > 0)
            {
                w = vw;
                h = vh;
            }
        }

        private static RasterizeResult Fail(string error)
        {
            return new RasterizeResult { Success = false, Error = error };
        }

        internal static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // ignore
            }
        }
    }
}
