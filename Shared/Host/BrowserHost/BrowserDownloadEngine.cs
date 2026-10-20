using System;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>url 直取 / ref+DownloadStarting → 会话工作区。</summary>
    internal static class BrowserDownloadEngine
    {
        public const long MaxBytes = 50L * 1024 * 1024;
        public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(30);
        public const string DefaultSubdir = "browser_dl";

        public static async Task<BrowserDownloadFileResult> FetchUrlAsync(
            YiWriteBrowserForm form,
            string url)
        {
            if (!Uri.TryCreate((url ?? "").Trim(), UriKind.Absolute, out Uri uri)
                || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("url 须为 http 或 https");
            }

            EnsureWorkspaceReady();

            using (var cts = new CancellationTokenSource(DefaultTimeout))
            using (var client = new HttpClient())
            {
                client.Timeout = DefaultTimeout;
                await TryAttachCookiesAsync(form, uri, client).ConfigureAwait(true);

                using (HttpResponseMessage resp = await client.GetAsync(
                    uri,
                    HttpCompletionOption.ResponseHeadersRead,
                    cts.Token).ConfigureAwait(true))
                {
                    if (!resp.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "下载失败 HTTP " + ((int)resp.StatusCode).ToString(CultureInfo.InvariantCulture));
                    }

                    if (resp.Content.Headers.ContentLength.HasValue
                        && resp.Content.Headers.ContentLength.Value > MaxBytes)
                    {
                        throw new InvalidOperationException("文件超过 50 MiB 上限");
                    }

                    string suggested = SuggestNameFromHeaders(
                        resp.Content.Headers.ContentDisposition?.FileName,
                        uri);
                    string relative = BuildRelativePath(suggested);
                    string localPath = WorkspacePathResolver.ResolveWritePath(relative);

                    using (Stream net = await resp.Content.ReadAsStreamAsync().ConfigureAwait(true))
                    using (var fs = new FileStream(localPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        await CopyWithLimitAsync(net, fs, MaxBytes, cts.Token).ConfigureAwait(true);
                    }

                    long bytes = new FileInfo(localPath).Length;
                    if (bytes > MaxBytes)
                    {
                        TryDelete(localPath);
                        throw new InvalidOperationException("文件超过 50 MiB 上限");
                    }

                    string contentType = resp.Content.Headers.ContentType?.MediaType;

                    return new BrowserDownloadFileResult
                    {
                        RelativePath = relative,
                        Bytes = bytes,
                        ContentType = contentType,
                        SourceUrl = uri.AbsoluteUri
                    };
                }
            }
        }

        public static async Task<BrowserDownloadFileResult> ClickAndCaptureAsync(
            YiWriteBrowserForm form,
            int backendNodeId,
            string sessionId = null,
            TimeSpan? timeout = null,
            string engine = null)
        {
            EnsureWorkspaceReady();

            TimeSpan wait = timeout ?? DefaultTimeout;
            Task<BrowserCapturedDownload> waitTask = form.BeginExpectDownloadAsync(wait);
            try
            {
                await BrowserInteractEngine
                    .ClickAsync(form, backendNodeId, sessionId, ConfigManager.GetBrowserClickMode(engine))
                    .ConfigureAwait(true);
            }
            catch
            {
                form.CancelExpectDownload();
                try
                {
                    await waitTask.ConfigureAwait(true);
                }
                catch
                {
                }

                throw;
            }

            BrowserCapturedDownload captured;
            try
            {
                captured = await waitTask.ConfigureAwait(true);
            }
            catch (TimeoutException)
            {
                throw new InvalidOperationException(
                    "未产生浏览器下载；若为图片或已知直链请改用 url 参数");
            }

            if (captured != null && captured.Interrupted)
            {
                if (!string.IsNullOrEmpty(captured.LocalPath)
                    && File.Exists(captured.LocalPath)
                    && new FileInfo(captured.LocalPath).Length > 0)
                {
                    return FinishCapturedFile(captured);
                }

                string interrupt = string.IsNullOrEmpty(captured.InterruptReason)
                    ? "下载被中断"
                    : ("下载被中断（" + captured.InterruptReason + "）");
                System.Diagnostics.Debug.WriteLine(
                    "[BrowserDownload] " + interrupt + " uri=" + (captured.Uri ?? ""));

                if (!string.IsNullOrWhiteSpace(captured.Uri)
                    && Uri.TryCreate(captured.Uri.Trim(), UriKind.Absolute, out Uri iu)
                    && (iu.Scheme == Uri.UriSchemeHttp || iu.Scheme == Uri.UriSchemeHttps))
                {
                    try
                    {
                        string suggested = Path.GetFileName(
                            (captured.RelativePath ?? "").Replace('\\', '/'));
                        return await FetchInPageAsync(form, iu.AbsoluteUri, suggested)
                            .ConfigureAwait(true);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(
                            interrupt + "；页内重取失败: " + (ex.Message ?? "未知错误"));
                    }
                }

                throw new InvalidOperationException(interrupt);
            }

            return FinishCapturedFile(captured);
        }

        /// <summary>
        /// 下载事件中断后，在页面里带 Cookie/Referer 再取，不走进程 HttpClient。
        /// </summary>
        public static async Task<BrowserDownloadFileResult> FetchInPageAsync(
            YiWriteBrowserForm form,
            string url,
            string suggestedFileName = null)
        {
            if (form == null)
            {
                throw new InvalidOperationException("引擎不可用");
            }

            if (!Uri.TryCreate((url ?? "").Trim(), UriKind.Absolute, out Uri uri)
                || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("页内重取须为 http 或 https");
            }

            EnsureWorkspaceReady();

            string urlJson = JsonConvert.SerializeObject(uri.AbsoluteUri);
            string expression =
                "(async function(){"
                + "const r=await fetch(" + urlJson + ",{credentials:'include',redirect:'follow'});"
                + "if(!r.ok)return{ok:false,status:r.status};"
                + "const buf=await r.arrayBuffer();"
                + "const n=buf.byteLength;"
                + "if(n>" + MaxBytes.ToString(CultureInfo.InvariantCulture) + ")return{ok:false,error:'too_large',bytes:n};"
                + "const bytes=new Uint8Array(buf);"
                + "let binary='';const chunk=0x8000;"
                + "for(let i=0;i<bytes.length;i+=chunk){binary+=String.fromCharCode.apply(null,bytes.subarray(i,i+chunk));}"
                + "return{ok:true,status:r.status,b64:btoa(binary),type:r.headers.get('content-type'),"
                + "disposition:r.headers.get('content-disposition'),bytes:n};"
                + "})()";

            var payload = new JObject
            {
                ["expression"] = expression,
                ["awaitPromise"] = true,
                ["returnByValue"] = true
            };

            string json = await form
                .CallCdpAsync("Runtime.evaluate", payload.ToString(Formatting.None))
                .ConfigureAwait(true);

            JObject root;
            try
            {
                root = JObject.Parse(json ?? "{}");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("页内重取返回无法解析: " + ex.Message);
            }

            JToken exn = root["exceptionDetails"];
            if (exn != null && exn.Type != JTokenType.Null)
            {
                string desc = exn["exception"]?["description"]?.ToString()
                    ?? exn["text"]?.ToString()
                    ?? "脚本异常";
                throw new InvalidOperationException(desc);
            }

            JToken value = root["result"]?["value"];
            if (value != null && value.Type == JTokenType.String)
            {
                try
                {
                    value = JToken.Parse(value.ToString());
                }
                catch
                {
                }
            }

            if (value == null || value.Type != JTokenType.Object)
            {
                throw new InvalidOperationException("页内重取无结果");
            }

            var body = (JObject)value;
            if (body["ok"] == null || body["ok"].Type != JTokenType.Boolean || !(bool)body["ok"])
            {
                int status = body["status"]?.Value<int>() ?? 0;
                string err = body["error"]?.ToString();
                if (string.Equals(err, "too_large", StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("文件超过 50 MiB 上限");
                }

                if (status > 0)
                {
                    throw new InvalidOperationException(
                        "页内重取失败 HTTP " + status.ToString(CultureInfo.InvariantCulture));
                }

                throw new InvalidOperationException("页内重取失败");
            }

            string b64 = body["b64"]?.ToString();
            if (string.IsNullOrEmpty(b64))
            {
                throw new InvalidOperationException("页内重取没有文件内容");
            }

            byte[] bytes;
            try
            {
                bytes = Convert.FromBase64String(b64);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("页内重取内容无法解码: " + ex.Message);
            }

            if (bytes == null || bytes.Length == 0)
            {
                throw new InvalidOperationException("页内重取文件为空");
            }

            if (bytes.Length > MaxBytes)
            {
                throw new InvalidOperationException("文件超过 50 MiB 上限");
            }

            if (LooksLikeHtml(bytes))
            {
                throw new InvalidOperationException("页内重取得到的是网页，不是文件");
            }

            string name = SuggestNameFromHeaders(body["disposition"]?.ToString(), uri);
            if (string.IsNullOrWhiteSpace(name))
            {
                name = suggestedFileName;
            }

            string relative = BuildRelativePath(name);
            string localPath = WorkspacePathResolver.ResolveWritePath(relative);
            File.WriteAllBytes(localPath, bytes);

            return new BrowserDownloadFileResult
            {
                RelativePath = relative,
                Bytes = bytes.Length,
                ContentType = body["type"]?.ToString(),
                SourceUrl = uri.AbsoluteUri
            };
        }

        private static BrowserDownloadFileResult FinishCapturedFile(BrowserCapturedDownload captured)
        {
            if (captured == null || string.IsNullOrEmpty(captured.LocalPath))
            {
                throw new InvalidOperationException("下载未完成");
            }

            if (!File.Exists(captured.LocalPath))
            {
                throw new InvalidOperationException("下载文件未落盘");
            }

            long nbytes = new FileInfo(captured.LocalPath).Length;
            if (nbytes > MaxBytes)
            {
                TryDelete(captured.LocalPath);
                throw new InvalidOperationException("文件超过 50 MiB 上限");
            }

            string relative = captured.RelativePath;
            if (string.IsNullOrEmpty(relative))
            {
                relative = BuildRelativePath(Path.GetFileName(captured.LocalPath));
            }

            return new BrowserDownloadFileResult
            {
                RelativePath = relative,
                Bytes = nbytes,
                ContentType = captured.ContentType,
                SourceUrl = captured.Uri
            };
        }

        private static bool LooksLikeHtml(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 12)
            {
                return false;
            }

            int i = 0;
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            {
                i = 3;
            }

            while (i < bytes.Length && i < 64 && (bytes[i] == (byte)' ' || bytes[i] == (byte)'\t'
                || bytes[i] == (byte)'\r' || bytes[i] == (byte)'\n'))
            {
                i++;
            }

            if (i + 5 > bytes.Length)
            {
                return false;
            }

            string head = Encoding.ASCII.GetString(bytes, i, Math.Min(15, bytes.Length - i))
                .ToLowerInvariant();
            return head.StartsWith("<!doctype", StringComparison.Ordinal)
                || head.StartsWith("<html", StringComparison.Ordinal);
        }

        public static string AllocateLocalPathForSuggestedName(string suggestedFileName, out string relativePath)
        {
            relativePath = BuildRelativePath(suggestedFileName);
            return WorkspacePathResolver.ResolveWritePath(relativePath);
        }

        internal static void EnsureWorkspaceReadyPublic() => EnsureWorkspaceReady();

        internal static string BuildRelativePathPublic(string suggestedFileName) => BuildRelativePath(suggestedFileName);

        internal static Task UploadOrThrowPublicAsync(string localPath, string relative) =>
            UploadOrThrowAsync(localPath, relative);

        internal static Task CopyWithLimitPublicAsync(
            Stream source,
            Stream dest,
            long maxBytes,
            CancellationToken ct) =>
            CopyWithLimitAsync(source, dest, maxBytes, ct);

        internal static string SuggestNameFromUriPublic(Uri uri) =>
            SuggestNameFromHeaders(null, uri) ?? "download.bin";

        private static void EnsureWorkspaceReady()
        {
            try
            {
                string dir = WorkspacePathResolver.GetSessionDirectory();
                if (string.IsNullOrWhiteSpace(dir))
                {
                    throw new InvalidOperationException("工作区未初始化，请先登录并配置工作目录");
                }
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("工作区未初始化: " + ex.Message);
            }
        }

        private static async Task TryAttachCookiesAsync(YiWriteBrowserForm form, Uri uri, HttpClient client)
        {
            if (form == null || !form.IsCoreReady)
            {
                return;
            }

            try
            {
                CoreWebView2Cookie[] cookies = await form.GetCookiesAsync(uri.AbsoluteUri).ConfigureAwait(true);
                if (cookies == null || cookies.Length == 0)
                {
                    return;
                }

                var sb = new StringBuilder();
                foreach (CoreWebView2Cookie c in cookies)
                {
                    if (c == null || string.IsNullOrEmpty(c.Name))
                    {
                        continue;
                    }

                    if (sb.Length > 0)
                    {
                        sb.Append("; ");
                    }

                    sb.Append(c.Name).Append('=').Append(c.Value ?? "");
                }

                if (sb.Length > 0)
                {
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Cookie", sb.ToString());
                }
            }
            catch
            {
                // cookie 可选
            }
        }

        private static string BuildRelativePath(string suggestedFileName)
        {
            string baseName = SanitizeLeaf(suggestedFileName);
            if (string.IsNullOrEmpty(baseName))
            {
                baseName = "download_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".bin";
            }

            string relative = DefaultSubdir + "/" + baseName;
            string sanitized = WorkspacePathResolver.SanitizeWorkspaceRelativePath(relative);
            if (sanitized == null)
            {
                throw new InvalidOperationException("无法生成合法工作区文件名");
            }

            // 重名加 (n)
            string sessionDir = WorkspacePathResolver.GetSessionDirectory();
            string candidate = sanitized;
            string dirPart = Path.GetDirectoryName(candidate.Replace('/', Path.DirectorySeparatorChar)) ?? "";
            string nameOnly = Path.GetFileNameWithoutExtension(candidate);
            string ext = Path.GetExtension(candidate);
            int n = 1;
            while (File.Exists(Path.Combine(sessionDir, candidate.Replace('/', Path.DirectorySeparatorChar))))
            {
                string leaf = nameOnly + "(" + n.ToString(CultureInfo.InvariantCulture) + ")" + ext;
                candidate = string.IsNullOrEmpty(dirPart)
                    ? leaf
                    : (dirPart.Replace('\\', '/') + "/" + leaf);
                candidate = WorkspacePathResolver.SanitizeWorkspaceRelativePath(candidate) ?? candidate;
                n++;
                if (n > 200)
                {
                    break;
                }
            }

            return candidate;
        }

        private static string SuggestNameFromHeaders(string contentDispositionFileName, Uri uri)
        {
            string fromCd = Unquote(contentDispositionFileName);
            if (!string.IsNullOrWhiteSpace(fromCd))
            {
                return Path.GetFileName(fromCd.Replace('\\', '/'));
            }

            try
            {
                string last = Path.GetFileName(uri.AbsolutePath);
                if (!string.IsNullOrWhiteSpace(last))
                {
                    return last;
                }
            }
            catch
            {
            }

            return null;
        }

        private static string Unquote(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
            {
                return null;
            }

            s = s.Trim();
            if (s.Length >= 2 && s[0] == '"' && s[s.Length - 1] == '"')
            {
                s = s.Substring(1, s.Length - 2);
            }

            return s;
        }

        private static string SanitizeLeaf(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            name = Path.GetFileName(name.Replace('\\', '/').Trim());
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            var sb = new StringBuilder(name.Length);
            foreach (char ch in name)
            {
                if (ch < 32 || "<>:\"/\\|?*".IndexOf(ch) >= 0)
                {
                    sb.Append('_');
                }
                else
                {
                    sb.Append(ch);
                }
            }

            string cleaned = sb.ToString().Trim('.', ' ');
            return string.IsNullOrEmpty(cleaned) ? null : cleaned;
        }

        private static async Task CopyWithLimitAsync(
            Stream source,
            Stream dest,
            long maxBytes,
            CancellationToken ct)
        {
            var buffer = new byte[81920];
            long total = 0;
            int read;
            while ((read = await source.ReadAsync(buffer, 0, buffer.Length, ct).ConfigureAwait(true)) > 0)
            {
                total += read;
                if (total > maxBytes)
                {
                    throw new InvalidOperationException("文件超过 50 MiB 上限");
                }

                await dest.WriteAsync(buffer, 0, read, ct).ConfigureAwait(true);
            }
        }

        private static Task UploadOrThrowAsync(string localPath, string relative)
        {
            _ = localPath;
            _ = relative;
            return Task.CompletedTask;
        }

        private static void TryDelete(string path)
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
            }
        }
    }

    internal sealed class BrowserDownloadFileResult
    {
        public string RelativePath { get; set; }
        public long Bytes { get; set; }
        public string ContentType { get; set; }
        public string SourceUrl { get; set; }
    }

    internal sealed class BrowserCapturedDownload
    {
        public string LocalPath { get; set; }
        public string RelativePath { get; set; }
        public string Uri { get; set; }
        public string ContentType { get; set; }
        public bool Interrupted { get; set; }
        public string InterruptReason { get; set; }
    }
}
