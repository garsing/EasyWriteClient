using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>attach 页的 CDP：经扩展 debugger。</summary>
    internal sealed class AttachCdpSession : IBrowserCdp
    {
        private readonly string _tabUuid;
        private string _url;
        private string _title;

        public AttachCdpSession(string tabUuid, string url, string title)
        {
            _tabUuid = tabUuid ?? "";
            _url = url ?? "";
            _title = title ?? "";
        }

        public string CurrentUrl
        {
            get { return _url; }
        }

        public string CurrentTitle
        {
            get { return _title; }
        }

        public void UpdateMeta(string url, string title)
        {
            if (!string.IsNullOrWhiteSpace(url))
            {
                _url = url;
            }

            if (title != null)
            {
                _title = title;
            }
        }

        public async Task RefreshMetaAsync()
        {
            try
            {
                string json = await CallCdpAsync(
                    "Runtime.evaluate",
                    "{\"expression\":\"({url:document.location.href,title:document.title})\",\"returnByValue\":true}")
                    .ConfigureAwait(true);
                JObject jo = JObject.Parse(json);
                JObject v = jo["result"]?["value"] as JObject;
                if (v != null)
                {
                    UpdateMeta((string)v["url"], (string)v["title"]);
                }
            }
            catch
            {
            }
        }

        public Task<string> CallCdpAsync(string methodName, string parametersAsJson, string sessionId = null)
        {
            return ExtensionHost.CallCdpAsync(_tabUuid, methodName, parametersAsJson, sessionId);
        }
    }
}
