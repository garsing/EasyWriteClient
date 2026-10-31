using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 浏览窗前面的确认框。WebView2 关掉默认脚本框后才把内容交给宿主，所以按原文和按钮自己画。
    /// 人点按钮，或智能体同意/拒绝，关的是同一扇窗。
    /// </summary>
    internal static class BrowserScriptDialogForm
    {
        public static bool ShowModal(Control host, BrowserScriptDialogSession session, string title)
        {
            if (host == null)
            {
                throw new ArgumentNullException(nameof(host));
            }

            if (session == null)
            {
                throw new ArgumentNullException(nameof(session));
            }

            Form owner = host.FindForm();
            using (var dlg = new Surface(session, title))
            {
                session.BindSurface(dlg, dlg.ApplyText, dlg.Finish);
                if (!session.IsPending)
                {
                    return session.Accepted;
                }

                DialogResult result = owner != null && !owner.IsDisposed
                    ? dlg.ShowDialog(owner)
                    : dlg.ShowDialog();
                return result == DialogResult.OK;
            }
        }

        private sealed class Surface : Form
        {
            private readonly BrowserScriptDialogSession _session;
            private readonly TextBox _input;

            public Surface(BrowserScriptDialogSession session, string title)
            {
                _session = session;
                Text = string.IsNullOrWhiteSpace(title) ? "网页消息" : title.Trim();
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;
                StartPosition = FormStartPosition.CenterParent;
                Font = SystemFonts.MessageBoxFont;

                string message = VisibleMessage(session);
                var label = new Label
                {
                    AutoSize = false,
                    Font = Font,
                    Location = new Point(16, 16),
                    MaximumSize = new Size(420, 0),
                    Text = message
                };
                Size preferred = label.GetPreferredSize(new Size(420, 0));
                label.Size = new Size(Math.Max(preferred.Width, 240), Math.Max(preferred.Height, 20));
                Controls.Add(label);

                int y = label.Bottom + 12;
                if (session.Kind == BrowserScriptDialogKind.Prompt)
                {
                    _input = new TextBox
                    {
                        Font = Font,
                        Location = new Point(16, y),
                        Width = Math.Max(label.Width, 280),
                        Text = session.Text ?? ""
                    };
                    _input.TextChanged += (s, e) => _session.NoteTypedByUser(_input.Text);
                    Controls.Add(_input);
                    y = _input.Bottom + 12;
                }

                bool alertOnly = session.Kind == BrowserScriptDialogKind.Alert;
                var ok = new Button
                {
                    Text = "确定",
                    DialogResult = DialogResult.OK,
                    Size = new Size(88, 28)
                };
                Controls.Add(ok);
                AcceptButton = ok;

                Button cancel = null;
                if (!alertOnly)
                {
                    cancel = new Button
                    {
                        Text = "取消",
                        DialogResult = DialogResult.Cancel,
                        Size = new Size(88, 28)
                    };
                    Controls.Add(cancel);
                    CancelButton = cancel;
                }

                int buttonsWidth = alertOnly ? ok.Width : ok.Width + 8 + cancel.Width;
                int contentWidth = label.Width;
                if (_input != null)
                {
                    contentWidth = Math.Max(contentWidth, _input.Width);
                }

                int clientWidth = Math.Max(contentWidth, buttonsWidth) + 32;
                int right = clientWidth - 16;
                if (cancel != null)
                {
                    cancel.Location = new Point(right - cancel.Width, y);
                    ok.Location = new Point(cancel.Left - 8 - ok.Width, y);
                }
                else
                {
                    ok.Location = new Point(right - ok.Width, y);
                }

                ClientSize = new Size(clientWidth, y + ok.Height + 16);
            }

            public void ApplyText(string text)
            {
                if (_input == null || _input.IsDisposed)
                {
                    return;
                }

                string value = text ?? "";
                if (string.Equals(_input.Text, value, StringComparison.Ordinal))
                {
                    return;
                }

                _input.Text = value;
                _input.SelectionStart = _input.Text.Length;
            }

            public void Finish(bool accept)
            {
                DialogResult = accept ? DialogResult.OK : DialogResult.Cancel;
            }

            private static string VisibleMessage(BrowserScriptDialogSession session)
            {
                if (!string.IsNullOrEmpty(session.Message))
                {
                    return session.Message;
                }

                if (session.Kind == BrowserScriptDialogKind.Beforeunload)
                {
                    return "是否离开此页面？系统可能不会保存您所做的更改。";
                }

                return "";
            }
        }
    }
}
