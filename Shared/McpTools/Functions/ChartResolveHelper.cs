using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 正文图表队列：内嵌图 + 浮动图。页眉、页脚、文本框不进。
    /// </summary>
    public static class ChartResolveHelper
    {
        public sealed class BodyChart
        {
            public int Start { get; set; }
            public Word.InlineShape Inline { get; set; }
            public Word.Shape Floating { get; set; }
            public Word.Chart Chart { get; set; }
            public object Raw { get; set; }

            public bool IsFloating
            {
                get { return Floating != null; }
            }

            public bool TryReadTitle(out string title)
            {
                title = "";
                try
                {
                    if (Inline != null)
                    {
                        title = Inline.Title ?? "";
                        return true;
                    }

                    if (Floating != null)
                    {
                        title = Floating.Title ?? "";
                        return true;
                    }

                    return WordAddIn1.ChartOoxmlHost.WordChartPackageHost.TryReadShapeTitle(Raw, out title);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[ChartStamp] 读 Title 失败: " + ex.Message);
                    return false;
                }
            }

            public bool TryWriteTitle(string title)
            {
                try
                {
                    if (Inline != null)
                    {
                        Inline.Title = title ?? "";
                        return true;
                    }

                    if (Floating != null)
                    {
                        Floating.Title = title ?? "";
                        return true;
                    }

                    return WordAddIn1.ChartOoxmlHost.WordChartPackageHost.TryWriteShapeTitle(Raw, title);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[ChartStamp] 写 Title 失败: " + ex.Message);
                    return false;
                }
            }

            public void Delete()
            {
                if (Inline != null)
                {
                    Inline.Delete();
                    return;
                }

                if (Floating != null)
                {
                    Floating.Delete();
                    return;
                }

                if (Raw != null)
                {
                    WordAddIn1.OpenFiles.WpsCom.Invoke(Raw, "Delete");
                }
            }

            public int End
            {
                get
                {
                    try
                    {
                        if (Inline != null && Inline.Range != null)
                        {
                            return Inline.Range.End;
                        }

                        if (Floating != null && Floating.Anchor != null)
                        {
                            return Floating.Anchor.End;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("[ChartResolve] 读 End 失败: " + ex.Message);
                    }

                    return Start;
                }
            }

            public Word.Range NavigateRange()
            {
                if (Inline != null)
                {
                    return Inline.Range;
                }

                return Floating != null ? Floating.Anchor : null;
            }
        }

        public static List<BodyChart> GetBodyChartsInOrder(Word.Document document)
        {
            var charts = new List<BodyChart>();
            if (document == null)
            {
                return charts;
            }

            if (WpsChartExtract.IsWpsDocument(document))
            {
                foreach (var item in WpsChartExtract.CollectShapes(document))
                {
                    charts.Add(new BodyChart
                    {
                        Start = item.Start,
                        Raw = item.Shape,
                        Inline = item.Shape as Word.InlineShape,
                        Floating = item.Shape as Word.Shape,
                        Chart = TryChartFromRaw(item.Shape)
                    });
                }

                charts.Sort((a, b) => a.Start.CompareTo(b.Start));
                return charts;
            }

            try
            {
                foreach (Word.InlineShape shape in document.InlineShapes)
                {
                    if (shape.Type != Word.WdInlineShapeType.wdInlineShapeChart)
                    {
                        continue;
                    }

                    if (!IsMainStory(shape.Range))
                    {
                        continue;
                    }

                    charts.Add(new BodyChart
                    {
                        Start = shape.Range.Start,
                        Inline = shape,
                        Chart = TryGetInlineChart(shape),
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ChartResolve] 收集内嵌图失败: " + ex.Message);
            }

            try
            {
                foreach (Word.Shape shape in document.Shapes)
                {
                    if (!IsFloatingChart(shape))
                    {
                        continue;
                    }

                    if (shape.Anchor == null || !IsMainStory(shape.Anchor))
                    {
                        continue;
                    }

                    charts.Add(new BodyChart
                    {
                        Start = shape.Anchor.Start,
                        Floating = shape,
                        Chart = TryGetFloatingChart(shape),
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ChartResolve] 收集浮动图失败: " + ex.Message);
            }

            charts.Sort((a, b) => a.Start.CompareTo(b.Start));
            return charts;
        }

        public static BodyChart ResolveBodyChartByChartId(Word.Document document, string chartId)
        {
            if (document == null || string.IsNullOrEmpty(chartId))
            {
                return null;
            }

            int chartOrderIndex = DocumentState.GetChartIndex(chartId);
            if (chartOrderIndex < 0)
            {
                return null;
            }

            List<BodyChart> allCharts = GetBodyChartsInOrder(document);
            if (chartOrderIndex >= allCharts.Count)
            {
                return null;
            }

            return allCharts[chartOrderIndex];
        }

        public static Word.InlineShape ResolveInlineShapeByChartId(Word.Document document, string chartId)
        {
            BodyChart chart = ResolveBodyChartByChartId(document, chartId);
            return chart != null ? chart.Inline : null;
        }

        private static bool IsMainStory(Word.Range range)
        {
            if (range == null)
            {
                return false;
            }

            try
            {
                return range.StoryType == Word.WdStoryType.wdMainTextStory;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ChartResolve] 读 StoryType 失败: " + ex.Message);
                return false;
            }
        }

        private static bool IsFloatingChart(Word.Shape shape)
        {
            if (shape == null)
            {
                return false;
            }

            try
            {
                if (shape.Type == MsoShapeType.msoChart)
                {
                    return true;
                }
            }
            catch
            {
            }

            try
            {
                return shape.HasChart == MsoTriState.msoTrue;
            }
            catch
            {
                return false;
            }
        }

        private static Word.Chart TryGetInlineChart(Word.InlineShape shape)
        {
            try
            {
                return shape.Chart;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ChartResolve] 内嵌 Chart 失败: " + ex.Message);
                return null;
            }
        }

        private static Word.Chart TryChartFromRaw(object shape)
        {
            if (shape == null)
            {
                return null;
            }

            try
            {
                Word.InlineShape inline = shape as Word.InlineShape;
                if (inline != null)
                {
                    return TryGetInlineChart(inline);
                }

                Word.Shape floating = shape as Word.Shape;
                if (floating != null)
                {
                    return TryGetFloatingChart(floating);
                }
            }
            catch
            {
            }

            return null;
        }

        private static Word.Chart TryGetFloatingChart(Word.Shape shape)
        {
            try
            {
                return shape.Chart;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ChartResolve] 浮动 Chart 失败: " + ex.Message);
                return null;
            }
        }
    }
}
