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

        /// <summary>读回属性断言；空则只对结构。</summary>
        public List<AttrCheck> Attrs { get; set; }
    }

    internal sealed class AttrCheck
    {
        public string Key { get; set; }

        public string Expected { get; set; }

        /// <summary>exact / hex / num / box / applyonly</summary>
        public string Mode { get; set; }

        public static AttrCheck Exact(string key, string expected)
        {
            return new AttrCheck { Key = key, Expected = expected, Mode = "exact" };
        }

        public static AttrCheck Hex(string key, string expected)
        {
            return new AttrCheck { Key = key, Expected = expected, Mode = "hex" };
        }

        public static AttrCheck Num(string key, string expected)
        {
            return new AttrCheck { Key = key, Expected = expected, Mode = "num" };
        }

        public static AttrCheck Box(string key, string expected)
        {
            return new AttrCheck { Key = key, Expected = expected, Mode = "box" };
        }

        public static AttrCheck ApplyOnly(string key, string expected)
        {
            return new AttrCheck { Key = key, Expected = expected, Mode = "applyonly" };
        }
    }

    internal sealed class ChartSeriesSpec
    {
        public string Type { get; set; }

        public string Axis { get; set; }

        public string Name { get; set; }

        public List<double> Values { get; set; }
    }
}
