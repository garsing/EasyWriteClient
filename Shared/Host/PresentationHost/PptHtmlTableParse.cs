using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace WordAddIn1.PresentationHost
{
    /// <summary>页内 table 约定 HTML 解析 / 拓扑预检（无 COM）。</summary>
    internal static partial class PptHtmlTableParse
    {
        private static readonly HashSet<string> SpanCssKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "color", "font-weight", "font-style"
        };

        public static bool TryParseTableShape(
            XElement tableEl,
            out PptHtmlTableGrid grid,
            out PptHtmlTableStyleSnap style,
            out string error,
            List<string> warnings)
        {
            grid = null;
            style = null;
            error = null;
            if (tableEl == null)
            {
                error = "缺少 table 元素";
                return false;
            }

            if (tableEl.Descendants().Any(e =>
                string.Equals(e.Name.LocalName, "table", StringComparison.OrdinalIgnoreCase)
                && !ReferenceEquals(e, tableEl)))
            {
                error = "不支持表格内再嵌套 <table>";
                return false;
            }

            if (!TryParseGrid(tableEl, forApply: true, out grid, out error, warnings))
            {
                return false;
            }

            style = ParseStyleAttrs(tableEl, grid.ColCount, grid.RowCount, warnings);
            return true;
        }

        public static bool TryParseGrid(
            XElement table,
            bool forApply,
            out PptHtmlTableGrid grid,
            out string error,
            List<string> warnings)
        {
            grid = null;
            error = null;
            if (warnings == null)
            {
                warnings = new List<string>();
            }

            List<XElement> rowEls = PptHtmlChartIo.EnumerateTableRows(table).ToList();
            if (rowEls.Count == 0)
            {
                error = "表格至少需要一行";
                return false;
            }

            var rawRows = new List<List<RawCell>>();
            int maxSlots = 0;
            foreach (XElement tr in rowEls)
            {
                var cells = new List<RawCell>();
                foreach (XElement td in tr.Elements().Where(e =>
                    string.Equals(e.Name.LocalName, "td", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(e.Name.LocalName, "th", StringComparison.OrdinalIgnoreCase)))
                {
                    if (!TryReadRawCell(td, out RawCell raw, out error, warnings))
                    {
                        return false;
                    }

                    cells.Add(raw);
                }

                if (cells.Count == 0)
                {
                    error = "表格行不能没有 <td>/<th>";
                    return false;
                }

                int slots = 0;
                foreach (RawCell c in cells)
                {
                    slots += Math.Max(1, c.ColSpan);
                }

                if (slots > maxSlots)
                {
                    maxSlots = slots;
                }

                rawRows.Add(cells);
            }

            int rows = rawRows.Count;
            int cols = maxSlots;
            if (forApply && (rows > PptHtmlTableGrid.MaxRows || cols > PptHtmlTableGrid.MaxCols))
            {
                error = "页内表格最多 " + PptHtmlTableGrid.MaxRows + " 行 × "
                    + PptHtmlTableGrid.MaxCols + " 列（含合并占位）";
                return false;
            }

            var occupied = new bool[rows, cols];
            var result = new PptHtmlTableCell[rows * cols];
            for (int i = 0; i < result.Length; i++)
            {
                result[i] = new PptHtmlTableCell { IsCovered = true, Text = "" };
            }

            for (int r = 0; r < rows; r++)
            {
                int cursor = 0;
                foreach (RawCell raw in rawRows[r])
                {
                    while (cursor < cols && occupied[r, cursor])
                    {
                        cursor++;
                    }

                    if (cursor >= cols)
                    {
                        error = "合并占位不合法（第 " + (r + 1) + " 行槽位溢出）";
                        return false;
                    }

                    int rs = Math.Max(1, raw.RowSpan);
                    int cs = Math.Max(1, raw.ColSpan);
                    if (r + rs > rows || cursor + cs > cols)
                    {
                        error = "合并占位不合法（越界 rowspan/colspan）";
                        return false;
                    }

                    for (int i = 0; i < rs; i++)
                    {
                        for (int j = 0; j < cs; j++)
                        {
                            if (occupied[r + i, cursor + j])
                            {
                                error = "合并占位不合法（重叠）";
                                return false;
                            }

                            occupied[r + i, cursor + j] = true;
                        }
                    }

                    result[r * cols + cursor] = new PptHtmlTableCell
                    {
                        Text = raw.Text ?? "",
                        RowSpan = rs,
                        ColSpan = cs,
                        IsCovered = false,
                        IsHeader = raw.IsHeader,
                        Fill = raw.Fill,
                        FontColor = raw.FontColor,
                        FontBold = raw.FontBold,
                        FontItalic = raw.FontItalic
                    };

                    for (int i = 0; i < rs; i++)
                    {
                        for (int j = 0; j < cs; j++)
                        {
                            if (i == 0 && j == 0)
                            {
                                continue;
                            }

                            result[(r + i) * cols + (cursor + j)] = new PptHtmlTableCell
                            {
                                IsCovered = true,
                                Text = "",
                                RowSpan = 1,
                                ColSpan = 1
                            };
                        }
                    }

                    cursor += cs;
                }

                for (int c = 0; c < cols; c++)
                {
                    if (!occupied[r, c])
                    {
                        error = "合并占位不合法（第 " + (r + 1) + " 行缺占位格）";
                        return false;
                    }
                }
            }

            bool truncated = false;
            if (!forApply && (rows > PptHtmlTableGrid.MaxRows || cols > PptHtmlTableGrid.MaxCols))
            {
                int keepR = Math.Min(rows, PptHtmlTableGrid.MaxRows);
                int keepC = Math.Min(cols, PptHtmlTableGrid.MaxCols);
                var kept = new List<PptHtmlTableCell>(keepR * keepC);
                for (int r = 0; r < keepR; r++)
                {
                    for (int c = 0; c < keepC; c++)
                    {
                        PptHtmlTableCell src = result[r * cols + c];
                        PptHtmlTableCell copy = CloneCell(src);
                        if (!copy.IsCovered)
                        {
                            if (r + copy.RowSpan > keepR)
                            {
                                copy.RowSpan = keepR - r;
                            }

                            if (c + copy.ColSpan > keepC)
                            {
                                copy.ColSpan = keepC - c;
                            }
                        }

                        kept.Add(copy);
                    }
                }

                grid = new PptHtmlTableGrid
                {
                    RowCount = keepR,
                    ColCount = keepC,
                    Cells = kept
                };
                truncated = true;
            }
            else
            {
                grid = new PptHtmlTableGrid
                {
                    RowCount = rows,
                    ColCount = cols,
                    Cells = new List<PptHtmlTableCell>(result)
                };
            }

            if (truncated)
            {
                // 调用方把 truncated 记到 style
            }

            return true;
        }

        public static bool GridWasTruncated(int rawRows, int rawCols)
        {
            return rawRows > PptHtmlTableGrid.MaxRows || rawCols > PptHtmlTableGrid.MaxCols;
        }

        public static bool TryRejectChartTableEnhancements(XElement table, out string error)
        {
            error = null;
            if (table == null)
            {
                return true;
            }

            foreach (XElement cell in table.Descendants().Where(e =>
                string.Equals(e.Name.LocalName, "td", StringComparison.OrdinalIgnoreCase)
                || string.Equals(e.Name.LocalName, "th", StringComparison.OrdinalIgnoreCase)))
            {
                string cs = Attr(cell, "colspan");
                string rs = Attr(cell, "rowspan");
                if ((!string.IsNullOrWhiteSpace(cs) && cs.Trim() != "1")
                    || (!string.IsNullOrWhiteSpace(rs) && rs.Trim() != "1")
                    || !string.IsNullOrWhiteSpace(Attr(cell, "data-fill"))
                    || cell.Elements().Any(e =>
                        string.Equals(e.Name.LocalName, "span", StringComparison.OrdinalIgnoreCase)))
                {
                    error = "chart 内嵌表不支持合并或单元格样式，请只写纯数据网格";
                    return false;
                }
            }

            return true;
        }

        public static PptHtmlTableStyleSnap ParseStyleAttrs(
            XElement tableEl,
            int colCount,
            int rowCount,
            List<string> warnings)
        {
            var snap = new PptHtmlTableStyleSnap();
            string fontName = Attr(tableEl, "data-table-font-name");
            if (!string.IsNullOrWhiteSpace(fontName))
            {
                snap.FontName = fontName.Trim();
            }

            string fontSize = Attr(tableEl, "data-table-font-size");
            if (!string.IsNullOrWhiteSpace(fontSize))
            {
                string s = fontSize.Trim();
                if (s.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
                {
                    s = s.Substring(0, s.Length - 2).Trim();
                }

                if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double pt)
                    && pt > 0 && pt <= 400)
                {
                    snap.FontSizePt = pt;
                }
                else if (warnings != null)
                {
                    warnings.Add("忽略无效 data-table-font-size=" + fontSize);
                }
            }

            string fontColor = Attr(tableEl, "data-table-font-color");
            if (!string.IsNullOrWhiteSpace(fontColor))
            {
                if (TryNormColor(fontColor, out string c))
                {
                    snap.FontColor = c;
                }
                else if (warnings != null)
                {
                    warnings.Add("忽略无效 data-table-font-color=" + fontColor);
                }
            }

            string tableStyle = Attr(tableEl, "data-table-style");
            if (!string.IsNullOrWhiteSpace(tableStyle))
            {
                if (string.Equals(tableStyle.Trim(), "three-line", StringComparison.OrdinalIgnoreCase))
                {
                    snap.TableStyle = "three-line";
                }
                else if (warnings != null)
                {
                    warnings.Add("忽略未知 data-table-style=" + tableStyle);
                }
            }

            string colW = Attr(tableEl, "data-col-widths");
            if (!string.IsNullOrWhiteSpace(colW))
            {
                if (TryParsePctList(colW, colCount, out float[] pcts, out string werr))
                {
                    snap.ColWidthPcts = pcts;
                }
                else if (warnings != null)
                {
                    warnings.Add("忽略 data-col-widths：" + werr);
                }
            }

            string rowH = Attr(tableEl, "data-row-heights");
            if (!string.IsNullOrWhiteSpace(rowH))
            {
                if (TryParsePctList(rowH, rowCount, out float[] pcts, out string herr))
                {
                    snap.RowHeightPcts = pcts;
                }
                else if (warnings != null)
                {
                    warnings.Add("忽略 data-row-heights：" + herr);
                }
            }

            return snap;
        }

        private static bool TryReadRawCell(
            XElement td,
            out RawCell raw,
            out string error,
            List<string> warnings)
        {
            raw = new RawCell
            {
                IsHeader = string.Equals(td.Name.LocalName, "th", StringComparison.OrdinalIgnoreCase),
                ColSpan = 1,
                RowSpan = 1
            };
            error = null;

            string cs = Attr(td, "colspan");
            if (!string.IsNullOrWhiteSpace(cs))
            {
                if (!int.TryParse(cs.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int n) || n < 1)
                {
                    error = "非法 colspan";
                    return false;
                }

                raw.ColSpan = n;
            }

            string rs = Attr(td, "rowspan");
            if (!string.IsNullOrWhiteSpace(rs))
            {
                if (!int.TryParse(rs.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int n) || n < 1)
                {
                    error = "非法 rowspan";
                    return false;
                }

                raw.RowSpan = n;
            }

            string fill = Attr(td, "data-fill");
            if (!string.IsNullOrWhiteSpace(fill))
            {
                if (string.Equals(fill.Trim(), "none", StringComparison.OrdinalIgnoreCase))
                {
                    raw.Fill = "none";
                }
                else if (TryNormColor(fill, out string fc))
                {
                    raw.Fill = fc;
                }
                else if (warnings != null)
                {
                    warnings.Add("忽略无效 data-fill=" + fill);
                }
            }

            List<XElement> spans = td.Elements()
                .Where(e => string.Equals(e.Name.LocalName, "span", StringComparison.OrdinalIgnoreCase))
                .ToList();
            List<XElement> otherEls = td.Elements()
                .Where(e => !string.Equals(e.Name.LocalName, "span", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (otherEls.Count > 0 && warnings != null)
            {
                warnings.Add("表格单元格含非 span 子标签，已按纯文本处理");
            }

            if (spans.Count > 1 && warnings != null)
            {
                warnings.Add("表格单元格多层 span，已取纯文本并丢样式");
            }

            if (spans.Count == 1 && otherEls.Count == 0)
            {
                XElement span = spans[0];
                raw.Text = string.Concat(span.DescendantNodes().OfType<XText>().Select(t => t.Value)).TrimEnd();
                ApplySpanStyle(Attr(span, "style"), raw, warnings);
            }
            else
            {
                raw.Text = string.Concat(td.DescendantNodes().OfType<XText>().Select(t => t.Value)).TrimEnd();
            }

            return true;
        }

        private static void ApplySpanStyle(string style, RawCell raw, List<string> warnings)
        {
            if (string.IsNullOrWhiteSpace(style))
            {
                return;
            }

            foreach (string part in style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int colon = part.IndexOf(':');
                if (colon <= 0)
                {
                    continue;
                }

                string key = part.Substring(0, colon).Trim();
                string val = part.Substring(colon + 1).Trim();
                if (!SpanCssKeys.Contains(key))
                {
                    if (warnings != null)
                    {
                        warnings.Add("忽略未知单元格 CSS: " + key);
                    }

                    continue;
                }

                if (string.Equals(key, "color", StringComparison.OrdinalIgnoreCase))
                {
                    if (TryNormColor(val, out string c))
                    {
                        raw.FontColor = c;
                    }
                    else if (warnings != null)
                    {
                        warnings.Add("忽略无效 color=" + val);
                    }
                }
                else if (string.Equals(key, "font-weight", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.Equals(val, "bold", StringComparison.OrdinalIgnoreCase) || val == "700")
                    {
                        raw.FontBold = true;
                    }
                    else if (string.Equals(val, "normal", StringComparison.OrdinalIgnoreCase) || val == "400")
                    {
                        raw.FontBold = false;
                    }
                }
                else if (string.Equals(key, "font-style", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.Equals(val, "italic", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(val, "oblique", StringComparison.OrdinalIgnoreCase))
                    {
                        raw.FontItalic = true;
                    }
                    else if (string.Equals(val, "normal", StringComparison.OrdinalIgnoreCase))
                    {
                        raw.FontItalic = false;
                    }
                }
            }
        }

        private static bool TryParsePctList(string raw, int expectCount, out float[] pcts, out string error)
        {
            pcts = null;
            error = null;
            string[] parts = raw.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != expectCount)
            {
                error = "个数须为 " + expectCount + "，实际 " + parts.Length;
                return false;
            }

            pcts = new float[expectCount];
            float sum = 0;
            for (int i = 0; i < parts.Length; i++)
            {
                string p = parts[i].Trim();
                if (p.EndsWith("%", StringComparison.Ordinal))
                {
                    p = p.Substring(0, p.Length - 1).Trim();
                }

                if (!float.TryParse(p, NumberStyles.Float, CultureInfo.InvariantCulture, out float v) || v < 0)
                {
                    error = "无效百分比: " + parts[i];
                    return false;
                }

                pcts[i] = v;
                sum += v;
            }

            if (sum <= 0)
            {
                error = "百分比加总须大于 0";
                return false;
            }

            if (Math.Abs(sum - 100f) > 0.5f)
            {
                for (int i = 0; i < pcts.Length; i++)
                {
                    pcts[i] = pcts[i] * 100f / sum;
                }
            }

            return true;
        }

        private static bool TryNormColor(string raw, out string color)
        {
            color = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return false;
            }

            string t = raw.Trim();
            if (t.StartsWith("#", StringComparison.Ordinal) && t.Length == 7)
            {
                color = t.ToUpperInvariant();
                return Regex.IsMatch(color, "^#[0-9A-F]{6}$");
            }

            if (t.StartsWith("#", StringComparison.Ordinal) && t.Length == 4)
            {
                color = ("#" + t[1] + t[1] + t[2] + t[2] + t[3] + t[3]).ToUpperInvariant();
                return true;
            }

            return false;
        }

        private static PptHtmlTableCell CloneCell(PptHtmlTableCell src)
        {
            if (src == null)
            {
                return new PptHtmlTableCell { IsCovered = true, Text = "" };
            }

            return new PptHtmlTableCell
            {
                Text = src.Text,
                RowSpan = src.RowSpan,
                ColSpan = src.ColSpan,
                IsCovered = src.IsCovered,
                IsHeader = src.IsHeader,
                Fill = src.Fill,
                FontColor = src.FontColor,
                FontBold = src.FontBold,
                FontItalic = src.FontItalic
            };
        }

        private static string Attr(XElement el, string name)
        {
            XAttribute a = el.Attribute(name);
            return a == null ? null : a.Value;
        }
    }
}
