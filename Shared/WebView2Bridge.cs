using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

namespace WordAddIn1
{
    /// <summary>
    /// WebView2 通信桥接类
    /// 用于 C# 和 Vue 之间的双向通信
    /// </summary>
    public class WebView2Bridge
    {
        private WebView2 webView2;
        private Dictionary<string, Func<object, Task<object>>> messageHandlers;

        public WebView2Bridge(WebView2 webView2)
        {
            this.webView2 = webView2;
            this.messageHandlers = new Dictionary<string, Func<object, Task<object>>>();
            
            // 监听来自 JavaScript 的消息
            webView2.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
        }

        /// <summary>
        /// 注册消息处理器
        /// </summary>
        public void RegisterHandler(string type, Func<object, Task<object>> handler)
        {
            messageHandlers[type] = handler;
        }

        /// <summary>
        /// 发送消息到 JavaScript
        /// </summary>
        public void SendToJavaScript(string type, object data = null)
        {
            try
            {
                var message = new
                {
                    type = type,
                    data = data,
                    timestamp = DateTimeOffset.Now.ToUnixTimeMilliseconds()
                };

                string json = JsonConvert.SerializeObject(message);

                // 确保在 UI 线程上执行
                if (webView2.InvokeRequired)
                {
                    webView2.Invoke(new Action(() =>
                    {
                        try
                        {
                            if (webView2.CoreWebView2 != null)
                            {
                                webView2.CoreWebView2.PostWebMessageAsJson(json);
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"[WebView2Bridge] CoreWebView2 为 null");
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[WebView2Bridge] UI 线程发送消息失败: {ex.Message}");
                        }
                    }));
                }
                else
                {
                    if (webView2.CoreWebView2 != null)
                    {
                        webView2.CoreWebView2.PostWebMessageAsJson(json);
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[WebView2Bridge] CoreWebView2 为 null");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WebView2Bridge] 发送消息失败: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[WebView2Bridge] 堆栈跟踪: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 处理来自 JavaScript 的消息
        /// </summary>
        private async void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string json = e.TryGetWebMessageAsString();
                if (string.IsNullOrEmpty(json))
                    return;

                var message = JsonConvert.DeserializeObject<WebViewMessage>(json);
                if (message == null || string.IsNullOrEmpty(message.type))
                    return;

                // 查找并执行对应的处理器
                if (messageHandlers.TryGetValue(message.type, out var handler))
                {
                    var result = await handler(message.data);
                    
                    // 如果需要返回结果，发送响应消息（包含原始消息类型以便前端识别）
                    if (result != null)
                    {
                        var response = new
                        {
                            type = "messageResponse",
                            originalType = message.type,
                            data = result,
                            timestamp = DateTimeOffset.Now.ToUnixTimeMilliseconds()
                        };
                        SendToJavaScript("messageResponse", response);
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[WebView2Bridge] 未找到消息处理器: {message.type}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WebView2Bridge] 处理消息失败: {ex.Message}");
            }
        }

        /// <summary>
        /// WebView 消息模型
        /// </summary>
        private class WebViewMessage
        {
            public string type { get; set; }
            public object data { get; set; }
            public long timestamp { get; set; }
        }
    }
}

