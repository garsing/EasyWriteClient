using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 在局部截图的盒子里给 DOM 编号，拼成一张图。易写与附着共用。
    /// </summary>
    internal static class BrowserNodeSheet
    {
        internal const int MaxNodes = 15;
        internal const string TooManyNote = "底下节点太多，可以缩小截图范围";

        public static async Task<BrowserCaptureResult> CaptureAsync(
            IBrowserCdp cdp,
            BrowserChannel channel,
            string refId,
            int? up)
        {
            BrowserLocalCapture.LocalPlate plate = await BrowserLocalCapture
                .CapturePlateAsync(cdp, channel, refId, up, collectNodes: true)
                .ConfigureAwait(true);
            if (plate.Error != null)
            {
                return BrowserCaptureResult.Fail(plate.Error);
            }

            List<NodeItem> items = ReadItems(plate.Items);
            if (items.Count == 0 || items.Count > MaxNodes)
            {
                string note = items.Count > MaxNodes ? TooManyNote : null;
                return BrowserLocalCapture.FinishImage(cdp, channel, plate.Jpeg, note);
            }

            string sessionId = string.IsNullOrWhiteSpace(plate.Entry.CdpSessionId)
                ? null
                : plate.Entry.CdpSessionId;
            foreach (NodeItem item in items)
            {
                int? backendId = null;
                if (!string.IsNullOrWhiteSpace(item.XPath))
                {
                    backendId = await BrowserInteractEngine
                        .BackendIdForXPathAsync(cdp, item.XPath, sessionId)
                        .ConfigureAwait(true);
                }

                string existing = BrowserRefStore.FindExisting(channel.ChannelId, backendId, item.Css);
                if (!string.IsNullOrEmpty(existing))
                {
                    item.RefId = existing;
                    continue;
                }

                if (!backendId.HasValue && string.IsNullOrWhiteSpace(item.Css))
                {
                    item.RefId = null;
                    continue;
                }

                var entry = new BrowserRefEntry
                {
                    BackendDomNodeId = backendId,
                    AttachCssPath = string.IsNullOrWhiteSpace(item.Css) ? null : item.Css,
                    XPath = string.IsNullOrWhiteSpace(item.XPath) ? null : item.XPath,
                    Name = item.Text,
                    Role = "dom",
                    FrameId = plate.Entry.FrameId,
                    CdpSessionId = plate.Entry.CdpSessionId,
                    AttachFrameId = plate.Entry.AttachFrameId
                };
                item.RefId = BrowserRefStore.Append(channel.ChannelId, entry);
            }

            items.RemoveAll(item => string.IsNullOrEmpty(item.RefId));
            if (items.Count == 0)
            {
                return BrowserLocalCapture.FinishImage(cdp, channel, plate.Jpeg, null);
            }

            byte[] sheet = Compose(plate.Jpeg, plate.CssW, plate.CssH, items);
            return BrowserLocalCapture.FinishImage(cdp, channel, sheet, null);
        }

        private static List<NodeItem> ReadItems(JArray raw)
        {
            var items = new List<NodeItem>();
            if (raw == null)
            {
                return items;
            }

            foreach (JToken token in raw)
            {
                JObject obj = token as JObject;
                if (obj == null)
                {
                    continue;
                }

                double w = obj.Value<double?>("w") ?? 0;
                double h = obj.Value<double?>("h") ?? 0;
                if (w <= 1 || h <= 1)
                {
                    continue;
                }

                items.Add(new NodeItem
                {
                    X = obj.Value<double?>("x") ?? 0,
                    Y = obj.Value<double?>("y") ?? 0,
                    W = w,
                    H = h,
                    Text = (obj.Value<string>("text") ?? "").Trim(),
                    Css = obj.Value<string>("css") ?? "",
                    XPath = obj.Value<string>("xpath") ?? ""
                });
            }

            return items;
        }

        private static byte[] Compose(byte[] jpeg, double cssW, double cssH, List<NodeItem> items)
        {
            using (var stream = new MemoryStream(jpeg))
            using (var source = new Bitmap(stream))
            {
                double scaleX = cssW > 1 ? source.Width / cssW : 1;
                double scaleY = cssH > 1 ? source.Height / cssH : scaleX;
                using (var overview = new Bitmap(source))
                using (var g = Graphics.FromImage(overview))
                {
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                    using (var font = LabelFont(11f, bold: true))
                    using (var textBrush = new SolidBrush(Color.FromArgb(192, 57, 43)))
                    using (var fill = new SolidBrush(Color.White))
                    using (var pen = new Pen(Color.FromArgb(192, 57, 43), 2f))
                    {
                        foreach (NodeItem item in items)
                        {
                            Rectangle box = PixelBox(item, scaleX, scaleY, source.Width, source.Height);
                            if (box.Width <= 0 || box.Height <= 0)
                            {
                                continue;
                            }

                            g.DrawRectangle(pen, box);
                            SizeF size = g.MeasureString(item.RefId, font);
                            int bw = (int)Math.Ceiling(size.Width) + 6;
                            int bh = (int)Math.Ceiling(size.Height) + 2;
                            int bx = Math.Max(0, Math.Min(box.X, overview.Width - bw));
                            int by = Math.Max(0, box.Y - bh);
                            if (by + bh > overview.Height)
                            {
                                by = Math.Max(0, overview.Height - bh);
                            }

                            g.FillRectangle(fill, bx, by, bw, bh);
                            g.DrawRectangle(pen, bx, by, bw, bh);
                            g.DrawString(item.RefId, font, textBrush, bx + 3, by);
                        }
                    }

                    List<SheetRow> rows = GroupRows(items);
                    return Pack(overview, source, scaleX, scaleY, rows);
                }
            }
        }

        private static byte[] Pack(
            Bitmap overview,
            Bitmap source,
            double scaleX,
            double scaleY,
            List<SheetRow> rows)
        {
            using (var font = LabelFont(14f, bold: true))
            using (var measure = Graphics.FromImage(overview))
            {
                int labelW = 48;
                foreach (SheetRow row in rows)
                {
                    foreach (NodeItem item in row.Items)
                    {
                        int textW = (int)Math.Ceiling(measure.MeasureString(RowLabel(item, row), font).Width);
                        labelW = Math.Max(labelW, textW + 16 + row.Depth * 16);
                    }
                }

                int rowGap = 8;
                int totalH = overview.Height + 16;
                foreach (SheetRow row in rows)
                {
                    foreach (NodeItem item in row.Items)
                    {
                        Rectangle box = PixelBox(item, scaleX, scaleY, source.Width, source.Height);
                        totalH += Math.Max(box.Height, 28) + rowGap;
                    }
                }

                int width = Math.Max(overview.Width, labelW + source.Width);
                using (var canvas = new Bitmap(Math.Max(1, width), Math.Max(1, totalH), PixelFormat.Format24bppRgb))
                using (var g = Graphics.FromImage(canvas))
                {
                    g.Clear(Color.White);
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                    g.DrawImage(overview, 0, 0);
                    using (var line = new Pen(Color.FromArgb(220, 223, 230), 2f))
                    {
                        g.DrawLine(line, 0, overview.Height + 6, canvas.Width, overview.Height + 6);
                    }

                    int y = overview.Height + 14;
                    using (var brush = new SolidBrush(Color.FromArgb(30, 30, 30)))
                    using (var alt = new SolidBrush(Color.FromArgb(250, 250, 252)))
                    {
                        int i = 0;
                        foreach (SheetRow row in rows)
                        {
                            foreach (NodeItem item in row.Items)
                            {
                                Rectangle box = PixelBox(item, scaleX, scaleY, source.Width, source.Height);
                                int rh = Math.Max(box.Height, 28);
                                if (i % 2 == 0)
                                {
                                    g.FillRectangle(alt, 0, y, canvas.Width, rh);
                                }

                                string label = RowLabel(item, row);
                                SizeF size = g.MeasureString(label, font);
                                g.DrawString(label, font, brush, 8 + row.Depth * 16, y + (rh - size.Height) / 2f);
                                if (box.Width > 0 && box.Height > 0)
                                {
                                    g.DrawImage(
                                        source,
                                        new Rectangle(labelW, y + (rh - box.Height) / 2, box.Width, box.Height),
                                        box,
                                        GraphicsUnit.Pixel);
                                }

                                y += rh + rowGap;
                                i++;
                            }
                        }
                    }

                    using (var ms = new MemoryStream())
                    {
                        canvas.Save(ms, ImageFormat.Jpeg);
                        return ms.ToArray();
                    }
                }
            }
        }

        private static string RowLabel(NodeItem item, SheetRow row)
        {
            string label = item.RefId;
            if (!string.IsNullOrEmpty(item.Text))
            {
                label += " " + item.Text;
            }

            if (!string.IsNullOrEmpty(row.ParentRef))
            {
                label += "（父 " + row.ParentRef + "）";
            }

            return label;
        }

        private static List<SheetRow> GroupRows(List<NodeItem> items)
        {
            var rows = new List<SheetRow>();
            foreach (NodeItem item in items)
            {
                SheetRow found = null;
                foreach (SheetRow row in rows)
                {
                    if (SameBand(row.Top, row.Bottom, item.Y, item.H))
                    {
                        found = row;
                        break;
                    }
                }

                if (found == null)
                {
                    found = new SheetRow();
                    rows.Add(found);
                }

                found.Items.Add(item);
                found.Top = Math.Min(found.Top, item.Y);
                found.Bottom = Math.Max(found.Bottom, item.Y + item.H);
                found.Left = Math.Min(found.Left, item.X);
            }

            foreach (SheetRow row in rows)
            {
                row.Items.Sort((a, b) => a.X.CompareTo(b.X));
            }

            for (int i = 0; i < rows.Count; i++)
            {
                for (int j = i - 1; j >= 0; j--)
                {
                    if (rows[j].Left <= rows[i].Left - 8)
                    {
                        rows[i].Parent = rows[j];
                        rows[i].ParentRef = rows[j].Items.Count == 0 ? null : rows[j].Items[0].RefId;
                        break;
                    }
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                int depth = 0;
                SheetRow parent = rows[i].Parent;
                while (parent != null && depth < 8)
                {
                    depth++;
                    parent = parent.Parent;
                }

                rows[i].Depth = depth;
            }

            return rows;
        }

        private static bool SameBand(double top, double bottom, double y, double h)
        {
            double overlap = Math.Min(bottom, y + h) - Math.Max(top, y);
            if (overlap <= 0)
            {
                return false;
            }

            double rowH = Math.Max(1, bottom - top);
            return overlap > h * 0.5 && overlap > rowH * 0.5;
        }

        private static Rectangle PixelBox(NodeItem item, double scaleX, double scaleY, int limitW, int limitH)
        {
            int x = (int)Math.Round(item.X * scaleX);
            int y = (int)Math.Round(item.Y * scaleY);
            int w = Math.Max(1, (int)Math.Round(item.W * scaleX));
            int h = Math.Max(1, (int)Math.Round(item.H * scaleY));
            if (x < 0)
            {
                w += x;
                x = 0;
            }

            if (y < 0)
            {
                h += y;
                y = 0;
            }

            if (x >= limitW || y >= limitH)
            {
                return Rectangle.Empty;
            }

            w = Math.Min(w, limitW - x);
            h = Math.Min(h, limitH - y);
            if (w <= 0 || h <= 0)
            {
                return Rectangle.Empty;
            }

            return new Rectangle(x, y, w, h);
        }

        private static Font LabelFont(float size, bool bold)
        {
            try
            {
                return new Font("Microsoft YaHei", size, bold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel);
            }
            catch
            {
                return new Font(FontFamily.GenericSansSerif, size, bold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel);
            }
        }

        private sealed class NodeItem
        {
            public double X;
            public double Y;
            public double W;
            public double H;
            public string Text;
            public string Css;
            public string XPath;
            public string RefId;
        }

        private sealed class SheetRow
        {
            public List<NodeItem> Items = new List<NodeItem>();
            public double Top = double.MaxValue;
            public double Bottom = double.MinValue;
            public double Left = double.MaxValue;
            public SheetRow Parent;
            public string ParentRef;
            public int Depth;
        }
    }
}
