using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace WordAddIn1.BrowserHost
{
    internal enum BrowserScriptDialogKind
    {
        Alert,
        Confirm,
        Prompt,
        Beforeunload
    }

    internal enum BrowserScriptDialogStep
    {
        NotPending,
        Reject,
        Typed,
        Accept,
        Dismiss
    }

    /// <summary>
    /// 一页上挂起的脚本对话框。同意/拒绝由调用方在 UI 线程上 Commit。
    /// </summary>
    internal sealed class BrowserScriptDialogSession
    {
        private readonly Action<string> _accept;
        private readonly Action _complete;
        private int _done;

        private BrowserScriptDialogSession(
            BrowserScriptDialogKind kind,
            string message,
            string text,
            Action<string> accept,
            Action complete)
        {
            Kind = kind;
            Message = message ?? "";
            Text = text ?? "";
            _accept = accept;
            _complete = complete;
        }

        public BrowserScriptDialogKind Kind { get; }

        public string Message { get; }

        public string Text { get; private set; }

        public bool IsPending => _done == 0;

        public bool Accepted { get; private set; }

        public bool Completed => _done != 0;

        public string ResultText { get; private set; }

        public static BrowserScriptDialogSession Create(
            BrowserScriptDialogKind kind,
            string message,
            string initialText,
            Action<string> accept,
            Action complete)
        {
            string text = kind == BrowserScriptDialogKind.Prompt ? (initialText ?? "") : "";
            return new BrowserScriptDialogSession(kind, message, text, accept, complete);
        }

        public void SetText(string text)
        {
            if (!IsPending)
            {
                return;
            }

            Text = text ?? "";
        }

        public void Accept()
        {
            if (Interlocked.Exchange(ref _done, 1) != 0)
            {
                return;
            }

            ResultText = Text ?? "";
            Accepted = true;
            try
            {
                if (_accept != null)
                {
                    _accept(ResultText);
                }
            }
            finally
            {
                SafeComplete();
            }
        }

        public void Dismiss()
        {
            if (Interlocked.Exchange(ref _done, 1) != 0)
            {
                return;
            }

            SafeComplete();
        }

        private void SafeComplete()
        {
            try
            {
                if (_complete != null)
                {
                    _complete();
                }
            }
            catch
            {
            }
        }
    }

    internal sealed class BrowserScriptDialogView
    {
        public BrowserScriptDialogView(string treeText, Dictionary<string, BrowserRefEntry> refs)
        {
            TreeText = treeText ?? "";
            Refs = refs ?? new Dictionary<string, BrowserRefEntry>(StringComparer.OrdinalIgnoreCase);
        }

        public string TreeText { get; }

        public Dictionary<string, BrowserRefEntry> Refs { get; }
    }

    /// <summary>脚本对话框的树文本与交互判断。不读无障碍树，不经过风险词。</summary>
    internal static class BrowserScriptDialog
    {
        public const string MustHandleFirst = "请先处理确认框";

        public const string ControlAccept = "accept";

        public const string ControlDismiss = "dismiss";

        public const string ControlInput = "input";

        public static string KindToken(BrowserScriptDialogKind kind)
        {
            switch (kind)
            {
                case BrowserScriptDialogKind.Alert:
                    return "alert";
                case BrowserScriptDialogKind.Prompt:
                    return "prompt";
                case BrowserScriptDialogKind.Beforeunload:
                    return "beforeunload";
                default:
                    return "confirm";
            }
        }

        public static BrowserScriptDialogView Render(BrowserScriptDialogSession session)
        {
            if (session == null)
            {
                throw new ArgumentNullException(nameof(session));
            }

            var refs = new Dictionary<string, BrowserRefEntry>(StringComparer.OrdinalIgnoreCase);
            var sb = new StringBuilder();
            sb.Append("- dialog ")
                .Append(KindToken(session.Kind))
                .Append(" \"")
                .Append(Escape(session.Message))
                .Append('"')
                .AppendLine();

            int n = 1;
            if (session.Kind == BrowserScriptDialogKind.Prompt)
            {
                string inputId = "e" + n;
                n++;
                refs[inputId] = Entry(ControlInput, session.Message, "textbox");
                sb.Append("  - textbox ref=")
                    .Append(inputId)
                    .Append(" \"")
                    .Append(Escape(session.Message))
                    .Append("\" value=\"")
                    .Append(Escape(session.Text))
                    .Append('"')
                    .AppendLine();
            }

            string acceptId = "e" + n;
            n++;
            refs[acceptId] = Entry(ControlAccept, "确定", "button");
            sb.Append("  - button ref=").Append(acceptId).Append(" \"确定\"").AppendLine();

            if (session.Kind != BrowserScriptDialogKind.Alert)
            {
                string dismissId = "e" + n;
                refs[dismissId] = Entry(ControlDismiss, "取消", "button");
                sb.Append("  - button ref=").Append(dismissId).Append(" \"取消\"").AppendLine();
            }

            return new BrowserScriptDialogView(sb.ToString(), refs);
        }

        public static BrowserScriptDialogStep Apply(
            BrowserScriptDialogSession session,
            string channelId,
            string action,
            string refId,
            string text,
            out string error)
        {
            error = null;
            if (session == null || !session.IsPending)
            {
                return BrowserScriptDialogStep.NotPending;
            }

            string act = (action ?? "").Trim().ToLowerInvariant();
            BrowserRefEntry entry = null;
            bool onDialog = !string.IsNullOrWhiteSpace(refId)
                && BrowserRefStore.TryGet(channelId, refId.Trim(), out entry)
                && entry != null
                && entry.IsScriptDialog;

            if (act == "click" && onDialog && entry.ScriptDialogControl == ControlAccept)
            {
                return BrowserScriptDialogStep.Accept;
            }

            if (act == "click" && onDialog && entry.ScriptDialogControl == ControlDismiss)
            {
                return BrowserScriptDialogStep.Dismiss;
            }

            if (act == "type" && onDialog && entry.ScriptDialogControl == ControlInput)
            {
                if (text == null)
                {
                    error = "type 必须提供 text";
                    return BrowserScriptDialogStep.Reject;
                }

                session.SetText(text);
                return BrowserScriptDialogStep.Typed;
            }

            error = MustHandleFirst;
            return BrowserScriptDialogStep.Reject;
        }

        public static void Commit(BrowserScriptDialogSession session, BrowserScriptDialogStep step)
        {
            if (session == null)
            {
                return;
            }

            if (step == BrowserScriptDialogStep.Accept)
            {
                session.Accept();
            }
            else if (step == BrowserScriptDialogStep.Dismiss)
            {
                session.Dismiss();
            }
        }

        private static BrowserRefEntry Entry(string control, string name, string role)
        {
            return new BrowserRefEntry
            {
                IsScriptDialog = true,
                ScriptDialogControl = control,
                Role = role,
                Name = name ?? "",
            };
        }

        private static string Escape(string s)
        {
            return (s ?? "")
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", " ")
                .Replace("\n", " ");
        }
    }
}
