using System;
using System.Collections.Generic;
using System.Text;
using WordAddIn1.OpenFiles;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1.PresentationHost
{
    /// <summary>
    /// 概览用的每页文字。只读已打开稿上的文本框、表格格子和图表标题，图片只打标记。
    /// </summary>
    internal static class PresentationSlidePlainText
    {
        public const string PictureMark = "这里有一张图";

        public static string CollectPowerPoint(PowerPoint.Slide slide)
        {
            return CollectPowerPoint(slide, null);
        }

        /// <summary>pictureLine 为空时仍写「这里有一张图」。返回空字符串表示这张图不出现。</summary>
        public static string CollectPowerPoint(
            PowerPoint.Slide slide,
            Func<PowerPoint.Shape, string> pictureLine)
        {
            if (slide == null)
            {
                return "";
            }

            var lines = new List<string>();
            try
            {
                PowerPoint.Shapes shapes = slide.Shapes;
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        AppendPowerPointShape(shapes[i], lines, pictureLine);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }

            return Join(lines);
        }

        public static string CollectWpp(object slide)
        {
            return CollectWpp(slide, null);
        }

        public static string CollectWpp(object slide, Func<object, string> pictureLine)
        {
            if (slide == null)
            {
                return "";
            }

            var lines = new List<string>();
            try
            {
                object shapes = WppCom.GetProperty(slide, "Shapes");
                int count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count"));
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        object shape = WppCom.Invoke(shapes, "Item", i);
                        AppendWppShape(shape, lines, pictureLine);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }

            return Join(lines);
        }

        private static void AppendPowerPointShape(
            PowerPoint.Shape shape,
            List<string> lines,
            Func<PowerPoint.Shape, string> pictureLine)
        {
            if (shape == null)
            {
                return;
            }

            Office.MsoShapeType type;
            try
            {
                type = shape.Type;
            }
            catch (Exception)
            {
                return;
            }

            if (type == Office.MsoShapeType.msoGroup)
            {
                try
                {
                    PowerPoint.GroupShapes items = shape.GroupItems;
                    int count = items.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            AppendPowerPointShape(items[i], lines, pictureLine);
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
                catch (Exception)
                {
                }

                return;
            }

            if (TryAppendPowerPointTable(shape, lines))
            {
                return;
            }

            if (TryAppendPowerPointChartTitle(shape, lines))
            {
                return;
            }

            if (type == Office.MsoShapeType.msoPicture
                || type == Office.MsoShapeType.msoLinkedPicture)
            {
                AddPictureLine(lines, pictureLine == null ? PictureMark : pictureLine(shape));
                return;
            }

            TryAppendPowerPointText(shape, lines);
        }

        private static bool TryAppendPowerPointTable(PowerPoint.Shape shape, List<string> lines)
        {
            bool hasTable;
            try
            {
                hasTable = shape.HasTable == Office.MsoTriState.msoTrue;
            }
            catch (Exception)
            {
                return false;
            }

            if (!hasTable)
            {
                return false;
            }

            try
            {
                PowerPoint.Table table = shape.Table;
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;
                for (int r = 1; r <= rows; r++)
                {
                    var cells = new List<string>();
                    for (int c = 1; c <= cols; c++)
                    {
                        try
                        {
                            string text = table.Cell(r, c).Shape.TextFrame.TextRange.Text;
                            cells.Add(NormalizeInline(text));
                        }
                        catch (Exception)
                        {
                            cells.Add("");
                        }
                    }

                    string row = string.Join(" ", cells).Trim();
                    if (row.Length > 0)
                    {
                        lines.Add(row);
                    }
                }
            }
            catch (Exception)
            {
            }

            return true;
        }

        private static bool TryAppendPowerPointChartTitle(PowerPoint.Shape shape, List<string> lines)
        {
            bool hasChart;
            try
            {
                hasChart = shape.HasChart == Office.MsoTriState.msoTrue;
            }
            catch (Exception)
            {
                return false;
            }

            if (!hasChart)
            {
                return false;
            }

            try
            {
                PowerPoint.Chart chart = shape.Chart;
                if (chart.HasTitle)
                {
                    AddText(lines, chart.ChartTitle.Text);
                }
            }
            catch (Exception)
            {
            }

            return true;
        }

        private static void TryAppendPowerPointText(PowerPoint.Shape shape, List<string> lines)
        {
            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return;
                }

                AddText(lines, shape.TextFrame.TextRange.Text);
            }
            catch (Exception)
            {
            }
        }

        private static void AppendWppShape(object shape, List<string> lines, Func<object, string> pictureLine)
        {
            if (shape == null)
            {
                return;
            }

            int type;
            try
            {
                type = Convert.ToInt32(WppCom.GetProperty(shape, "Type"));
            }
            catch (Exception)
            {
                return;
            }

            if (type == (int)Office.MsoShapeType.msoGroup)
            {
                try
                {
                    object items = WppCom.GetProperty(shape, "GroupItems");
                    int count = Convert.ToInt32(WppCom.GetProperty(items, "Count"));
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            AppendWppShape(WppCom.Invoke(items, "Item", i), lines, pictureLine);
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
                catch (Exception)
                {
                }

                return;
            }

            if (IsTruthy(WppCom.GetProperty(shape, "HasTable")))
            {
                AppendWppTable(shape, lines);
                return;
            }

            if (IsTruthy(WppCom.GetProperty(shape, "HasChart")))
            {
                AppendWppChartTitle(shape, lines);
                return;
            }

            if (type == (int)Office.MsoShapeType.msoPicture
                || type == (int)Office.MsoShapeType.msoLinkedPicture)
            {
                AddPictureLine(lines, pictureLine == null ? PictureMark : pictureLine(shape));
                return;
            }

            try
            {
                if (!IsTruthy(WppCom.GetProperty(shape, "HasTextFrame")))
                {
                    return;
                }

                object frame = WppCom.GetProperty(shape, "TextFrame");
                object range = frame == null ? null : WppCom.GetProperty(frame, "TextRange");
                AddText(lines, range == null ? null : Convert.ToString(WppCom.GetProperty(range, "Text")));
            }
            catch (Exception)
            {
            }
        }

        private static void AppendWppTable(object shape, List<string> lines)
        {
            try
            {
                object table = WppCom.GetProperty(shape, "Table");
                object rows = WppCom.GetProperty(table, "Rows");
                object cols = WppCom.GetProperty(table, "Columns");
                int rowCount = Convert.ToInt32(WppCom.GetProperty(rows, "Count"));
                int colCount = Convert.ToInt32(WppCom.GetProperty(cols, "Count"));
                for (int r = 1; r <= rowCount; r++)
                {
                    var cells = new List<string>();
                    for (int c = 1; c <= colCount; c++)
                    {
                        try
                        {
                            object cell = WppCom.Invoke(table, "Cell", r, c);
                            object cellShape = WppCom.GetProperty(cell, "Shape");
                            object frame = WppCom.GetProperty(cellShape, "TextFrame");
                            object range = frame == null ? null : WppCom.GetProperty(frame, "TextRange");
                            cells.Add(NormalizeInline(range == null
                                ? null
                                : Convert.ToString(WppCom.GetProperty(range, "Text"))));
                        }
                        catch (Exception)
                        {
                            cells.Add("");
                        }
                    }

                    string row = string.Join(" ", cells).Trim();
                    if (row.Length > 0)
                    {
                        lines.Add(row);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void AppendWppChartTitle(object shape, List<string> lines)
        {
            try
            {
                object chart = WppCom.GetProperty(shape, "Chart");
                if (!IsTruthy(WppCom.GetProperty(chart, "HasTitle")))
                {
                    return;
                }

                object title = WppCom.GetProperty(chart, "ChartTitle");
                AddText(lines, title == null ? null : Convert.ToString(WppCom.GetProperty(title, "Text")));
            }
            catch (Exception)
            {
            }
        }

        private static void AddPictureLine(List<string> lines, string line)
        {
            if (!string.IsNullOrEmpty(line))
            {
                lines.Add(line);
            }
        }

        private static void AddText(List<string> lines, string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n').Trim();
            if (normalized.Length > 0)
            {
                lines.Add(normalized);
            }
        }

        private static string NormalizeInline(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return "";
            }

            return text.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ').Trim();
        }

        private static string Join(List<string> lines)
        {
            if (lines == null || lines.Count == 0)
            {
                return "";
            }

            var sb = new StringBuilder();
            for (int i = 0; i < lines.Count; i++)
            {
                if (i > 0)
                {
                    sb.Append('\n');
                }

                sb.Append(lines[i]);
            }

            return sb.ToString();
        }

        private static bool IsTruthy(object value)
        {
            if (value == null)
            {
                return false;
            }

            if (value is bool bit)
            {
                return bit;
            }

            try
            {
                return Convert.ToInt32(value) != 0;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
