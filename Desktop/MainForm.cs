using System;
using System.Drawing;
using System.Windows.Forms;
using WordAddIn1;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// B0 桌面壳：窗口标题「易写」。后续 B4 嵌入 WebView2（任务侧栏 + 对话框）。
    /// </summary>
    public sealed class MainForm : Form
    {
        private readonly Label _hintLabel;
        private readonly Label _channelLabel;

        public MainForm()
        {
            Text = AppDisplayName.Value;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(960, 640);
            Size = new Size(1280, 800);
            BackColor = Color.White;

            _hintLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei UI", 14F, FontStyle.Regular),
                ForeColor = Color.FromArgb(60, 60, 60),
                Text = "易写 Desktop（B0 骨架）\r\n后续将在此嵌入 WebView2：左侧任务 + 中间对话",
            };

            _channelLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Bottom,
                Height = 36,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(16, 0, 0, 0),
                Font = new Font("Microsoft YaHei UI", 9F),
                ForeColor = Color.FromArgb(100, 100, 100),
            };

            Controls.Add(_hintLabel);
            Controls.Add(_channelLabel);

            Shown += OnShown;
        }

        private void OnShown(object sender, EventArgs e)
        {
            // 冒烟：Shared 渠道注册表已链接
            string defaultId = ChannelRegistry.DefaultChannelId;
            _channelLabel.Text = string.IsNullOrEmpty(defaultId)
                ? "Shared OK · 默认渠道：无（请先打开文档）"
                : "Shared OK · 默认渠道：" + defaultId;
        }
    }
}
