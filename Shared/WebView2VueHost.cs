using System;
using Microsoft.Web.WebView2.Core;
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

            AttachImageContextMenuFilter(webView);
        }

        /// <summary>
        /// 图片右键只保留「复制图像」「另存为」，去掉在新标签打开、复制地址、检查等。
        /// </summary>
        public static void AttachImageContextMenuFilter(WebView2 webView)
        {
            if (webView?.CoreWebView2 == null)
            {
                throw new InvalidOperationException("CoreWebView2 尚未初始化。");
            }

            webView.CoreWebView2.ContextMenuRequested += (_, e) =>
            {
                try
                {
                    if (e.ContextMenuTarget == null ||
                        e.ContextMenuTarget.Kind != CoreWebView2ContextMenuTargetKind.Image)
                    {
                        return;
                    }

                    var items = e.MenuItems;
                    for (int i = items.Count - 1; i >= 0; i--)
                    {
                        if (!KeepImageContextMenuItem(items[i]))
                        {
                            items.RemoveAt(i);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        "[WebView2VueHost] ContextMenuRequested: " + ex.Message);
                }
            };
        }

        private static bool KeepImageContextMenuItem(CoreWebView2ContextMenuItem item)
        {
            if (item == null || item.Kind == CoreWebView2ContextMenuItemKind.Separator)
            {
                return false;
            }

            string name = item.Name ?? string.Empty;
            return string.Equals(name, "copyImage", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "saveAs", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "saveImageAs", StringComparison.OrdinalIgnoreCase);
        }
    }
}
