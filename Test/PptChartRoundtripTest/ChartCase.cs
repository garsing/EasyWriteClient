using System.Collections.Generic;

namespace PptChartRoundtripTest
{
    internal sealed class ChartCase
    {
        public int Index { get; set; }

        public int Batch { get; set; }

        public int Page { get; set; }

        public string Name { get; set; }

        public bool IsReplace { get; set; }

        public string CreateHtml { get; set; }

        public string ReplaceHtml { get; set; }

        public string ExpectType { get; set; }

        public int ExpectRows { get; set; }

        public int ExpectSeries { get; set; }

        public string ExpectLineAxis { get; set; }

        public string FirstCategory { get; set; }

        public double FirstValue { get; set; }

        public double LastValue { get; set; }
    }

    internal sealed class ChartSeriesSpec
    {
        public string Type { get; set; }

        public string Axis { get; set; }

        public string Name { get; set; }

        public List<double> Values { get; set; }
    }
}
