using System.Threading.Tasks;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 指针树 / CDP 操作的传输口。WebView2 与扩展 debugger 共用。
    /// </summary>
    internal interface IBrowserCdp
    {
        string CurrentUrl { get; }

        string CurrentTitle { get; }

        Task<string> CallCdpAsync(string methodName, string parametersAsJson, string sessionId = null);
    }
}
