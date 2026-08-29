using System;
using Microsoft.Web.WebView2.WinForms;

namespace WordAddIn1
{
    /// <summary>
    /// Vue 宿主 WebView2 的统一设置：禁止 Ctrl+滚轮 / 捏合缩放页面。
    /// </summary>
    public static class WebView2VueHost
    {
        /// <summary>
        /// 关闭用户缩放，并把倍率钉在 1.0。须在 EnsureCoreWebView2 之后调用。
        /// </summary>
        public static void DisableUserZoom(WebView2 webView)
        {
            if (webView == null)
            {
                throw new ArgumentNullException(nameof(webView));
            }

            if (webView.CoreWebView2 == null)
            {
                throw new InvalidOperationException("CoreWebView2 尚未初始化。");
            }

            var settings = webView.CoreWebView2.Settings;
            settings.IsZoomControlEnabled = false;
            settings.IsPinchZoomEnabled = false;
            webView.ZoomFactor = 1.0;
            webView.ZoomFactorChanged += (_, __) =>
            {
                if (Math.Abs(webView.ZoomFactor - 1.0) > 0.001)
                {
                    webView.ZoomFactor = 1.0;
                }
            };
        }
    }
}
