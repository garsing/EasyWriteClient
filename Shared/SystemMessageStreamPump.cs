using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAddIn1
{
    /// <summary>
    /// 流式 systemMessage 泵：不在 SSE 线程上同步 Invoke，
    /// 避免 UI 被堵死、多 chunk 攒完才一次性画出（加载点落在工具卡后）。
    /// 正文约 16ms 一帧；tool_calls / finish_reason 合并进同一次刷新且不丢增量。
    /// </summary>
    internal sealed class SystemMessageStreamPump
    {
        private const int PaceMs = 16;

        private readonly Control _ui;
        private readonly Action<string, string, object, string, bool> _send;
        private readonly object _gate = new object();
        private string _content = "";
        private string _reasoning;
        private readonly List<object> _toolItems = new List<object>();
        private string _finishReason;
        private bool _isFirst = true;
        private bool _queued;
        private string _lastSentContent = "";
        private bool _disposed;

        public SystemMessageStreamPump(Control ui, Action<string, string, object, string, bool> send)
        {
            _ui = ui ?? throw new ArgumentNullException(nameof(ui));
            _send = send ?? throw new ArgumentNullException(nameof(send));
        }

        public void Enqueue(string content, string reasoning, object toolCallsDelta, string finishReason)
        {
            if (_disposed)
            {
                return;
            }

            bool schedule;
            lock (_gate)
            {
                _content = content ?? "";
                _reasoning = reasoning;
                AppendToolDelta(toolCallsDelta);
                if (!string.IsNullOrEmpty(finishReason))
                {
                    _finishReason = finishReason;
                }

                if (_queued)
                {
                    return;
                }

                _queued = true;
                schedule = true;
            }

            if (schedule)
            {
                BeginFlush();
            }
        }

        public void Dispose()
        {
            _disposed = true;
        }

        private void BeginFlush()
        {
            if (_disposed || !_ui.IsHandleCreated)
            {
                return;
            }

            try
            {
                _ui.BeginInvoke(new Action(FlushCore));
            }
            catch (Exception)
            {
                lock (_gate)
                {
                    _queued = false;
                }
            }
        }

        private async void FlushCore()
        {
            if (_disposed)
            {
                return;
            }

            string content;
            string reasoning;
            object tools;
            string finish;
            bool first;
            lock (_gate)
            {
                content = _content;
                reasoning = _reasoning;
                tools = _toolItems.Count > 0 ? new List<object>(_toolItems) : null;
                _toolItems.Clear();
                finish = _finishReason;
                _finishReason = null;
                first = _isFirst;
                _isFirst = false;
                _lastSentContent = content;
            }

            try
            {
                _send(content, reasoning, tools, finish, first);
            }
            catch (Exception)
            {
            }

            try
            {
                await Task.Delay(PaceMs).ConfigureAwait(true);
            }
            catch (Exception)
            {
            }

            bool again;
            lock (_gate)
            {
                again = !_disposed && (
                    _toolItems.Count > 0
                    || !string.IsNullOrEmpty(_finishReason)
                    || !string.Equals(_content, _lastSentContent, StringComparison.Ordinal));
                if (!again)
                {
                    _queued = false;
                }
            }

            if (again)
            {
                BeginFlush();
            }
        }

        private void AppendToolDelta(object toolCallsDelta)
        {
            if (toolCallsDelta == null)
            {
                return;
            }

            if (toolCallsDelta is System.Collections.IEnumerable items
                && !(toolCallsDelta is string))
            {
                foreach (object item in items)
                {
                    if (item != null)
                    {
                        _toolItems.Add(item);
                    }
                }

                return;
            }

            _toolItems.Add(toolCallsDelta);
        }
    }
}
