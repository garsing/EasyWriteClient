using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public sealed class StructureValidationResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
    }

    public static class ChartStructureValidator
    {
        public static StructureValidationResult Validate(
            Word.Chart chart,
            ChartInfo parsedChart,
            string scope)
        {
            var result = new StructureValidationResult();
            if (chart == null)
            {
                result.Error = "图表对象不可用";
                return result;
            }

            string normalizedScope = string.IsNullOrEmpty(scope) ? "all" : scope.Trim().ToLowerInvariant();
            int existingSeries = GetSeriesCount(chart);
            int existingPoints = GetFirstSeriesPointCount(chart);

            if (normalizedScope == "format")
            {
                if (parsedChart?.SeriesCollection != null && parsedChart.SeriesCollection.Count > 0)
                {
                    int formatScopeSeriesCount = parsedChart.SeriesCollection.Count;
                    if (formatScopeSeriesCount != existingSeries)
                    {
                        result.Error = BuildMismatchMessage(existingSeries, formatScopeSeriesCount, existingPoints, 0);
                        return result;
                    }
                }

                result.Success = true;
                return result;
            }

            if (parsedChart?.SeriesCollection == null || parsedChart.SeriesCollection.Count == 0)
            {
                result.Error = "套用数据 scope 需要 DataTable 系列数据";
                return result;
            }

            int parsedSeries = parsedChart.SeriesCollection.Count;
            int parsedPoints = parsedChart.SeriesCollection[0].Points?.Count ?? 0;

            if (parsedSeries != existingSeries || parsedPoints != existingPoints)
            {
                result.Error = BuildMismatchMessage(existingSeries, parsedSeries, existingPoints, parsedPoints);
                return result;
            }

            result.Success = true;
            return result;
        }

        private static string BuildMismatchMessage(
            int expectedSeries,
            int actualSeries,
            int expectedPoints,
            int actualPoints)
        {
            return $"chart_data_structure_mismatch: 期望系列数={expectedSeries} 点数={expectedPoints}，"
                + $"XML 系列数={actualSeries} 点数={actualPoints}";
        }

        private static int GetSeriesCount(Word.Chart chart)
        {
            try
            {
                return chart.SeriesCollection().Count;
            }
            catch
            {
                return 0;
            }
        }

        private static int GetFirstSeriesPointCount(Word.Chart chart)
        {
            try
            {
                dynamic series = chart.SeriesCollection(1);
                try
                {
                    return series.Points().Count;
                }
                catch
                {
                    // Points().Count 不可用时再读 Values
                }

                object valuesObj = null;
                try
                {
                    var seriesObj = (object)series;
                    var t = seriesObj.GetType();
                    valuesObj = t.InvokeMember(
                        "Values",
                        System.Reflection.BindingFlags.GetProperty
                            | System.Reflection.BindingFlags.Public
                            | System.Reflection.BindingFlags.Instance,
                        null,
                        seriesObj,
                        null);
                }
                catch { }

                if (valuesObj is Array valuesArray)
                {
                    return valuesArray.Length;
                }
            }
            catch { }

            return 0;
        }
    }
}
