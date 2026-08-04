using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public sealed class ChartApplyResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public List<string> Warnings { get; set; } = new List<string>();
        public int SeriesUpdated { get; set; }
        public int PointsUpdated { get; set; }
    }

    public static class ChartApplyHelper
    {
        public static ChartApplyResult Apply(
            Word.Chart chart,
            Word.InlineShape inlineShape,
            ChartInfo chartInfo,
            string scope,
            string mode)
        {
            string normalizedMode = string.IsNullOrEmpty(mode) ? "full" : mode.Trim().ToLowerInvariant();
            if (normalizedMode != "full")
            {
                return new ChartApplyResult { Success = false, Error = $"mode 无效: '{mode}'，当前仅支持 full" };
            }

            return ChartEngine.ApplyToExistingChart(chart, inlineShape, chartInfo, scope);
        }
    }
}
