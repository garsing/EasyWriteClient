using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace WordAddIn1
{
    /// <summary>
    /// 人看输出汇合：Python SSE 与终端本机都走这里，再发同一种 toolOutput。
    /// </summary>
    public sealed class ToolOutputBridge : IDisposable
    {
        private const int FlushMs = 50;
        private const int FlushChars = 2048;

        public static ToolOutputBridge Current { get; private set; }

        private readonly WebView2Bridge _bridge;
        private readonly object _gate = new object();
        private readonly Dictionary<string, StringBuilder> _bufs = new Dictionary<string, StringBuilder>();
        private readonly Dictionary<string, DateTime> _started = new Dictionary<string, DateTime>();
        private Timer _timer;
        private bool _disposed;

        public ToolOutputBridge(WebView2Bridge bridge)
        {
            _bridge = bridge ?? throw new ArgumentNullException(nameof(bridge));
            Current = this;
        }

        public bool TryHandleMeta(EasyWriteStreamMeta meta)
        {
            if (meta == null)
            {
                return false;
            }

            if (meta.@event == "tool_output")
            {
                Emit(meta.tool_call_id, meta.text);
                return true;
            }

            if (meta.@event == "tool_output_done")
            {
                EmitDone(meta.tool_call_id);
                return true;
            }

            return false;
        }

        public void EmitDone(string toolCallId)
        {
            if (_disposed || string.IsNullOrEmpty(toolCallId) || _bridge == null)
            {
                return;
            }

            lock (_gate)
            {
                FlushOne(toolCallId);
                _bridge.SendToJavaScript("toolOutputDone", new
                {
                    tool_call_id = toolCallId
                });
            }
        }

        public void Emit(string toolCallId, string text)
        {
            if (_disposed || string.IsNullOrEmpty(toolCallId) || string.IsNullOrEmpty(text))
            {
                return;
            }

            lock (_gate)
            {
                if (!_bufs.TryGetValue(toolCallId, out StringBuilder sb))
                {
                    sb = new StringBuilder();
                    _bufs[toolCallId] = sb;
                    _started[toolCallId] = DateTime.UtcNow;
                }

                sb.Append(text);
                if (sb.Length >= FlushChars)
                {
                    FlushOne(toolCallId);
                }
                else
                {
                    EnsureTimer();
                }
            }
        }

        public void Dispose()
        {
            lock (_gate)
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                FlushAll();
                _timer?.Dispose();
                _timer = null;
                if (ReferenceEquals(Current, this))
                {
                    Current = null;
                }
            }
        }

        private void EnsureTimer()
        {
            if (_timer != null)
            {
                return;
            }

            _timer = new Timer(_ => FlushDue(), null, FlushMs, FlushMs);
        }

        private void FlushDue()
        {
            lock (_gate)
            {
                if (_disposed)
                {
                    return;
                }

                DateTime now = DateTime.UtcNow;
                var due = new List<string>();
                foreach (KeyValuePair<string, DateTime> pair in _started)
                {
                    if ((now - pair.Value).TotalMilliseconds >= FlushMs)
                    {
                        due.Add(pair.Key);
                    }
                }

                foreach (string id in due)
                {
                    FlushOne(id);
                }

                if (_bufs.Count == 0)
                {
                    _timer?.Dispose();
                    _timer = null;
                }
            }
        }

        private void FlushAll()
        {
            var ids = new List<string>(_bufs.Keys);
            foreach (string id in ids)
            {
                FlushOne(id);
            }
        }

        private void FlushOne(string toolCallId)
        {
            if (!_bufs.TryGetValue(toolCallId, out StringBuilder sb))
            {
                return;
            }

            string text = sb.ToString();
            _bufs.Remove(toolCallId);
            _started.Remove(toolCallId);
            if (string.IsNullOrEmpty(text) || _bridge == null)
            {
                return;
            }

            _bridge.SendToJavaScript("toolOutput", new
            {
                tool_call_id = toolCallId,
                text
            });
        }
    }
}
