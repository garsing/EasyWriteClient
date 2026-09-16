using System.Collections.Generic;

namespace PptChartRoundtripTest
{
    internal static class ChartCaseCatalog
    {
        public const int BatchSize = 50;

        public const int BatchCount = 6;

        public const int Total = BatchSize * BatchCount;

        private static readonly string[] Types = { "column", "bar", "line", "pie2d", "pie3d" };

        private static readonly string[] Legends = { "none", "top", "bottom", "left", "right" };

        private static readonly int[] RowCounts = { 4, 5, 6 };

        public static List<ChartCase> Build()
        {
            var list = new List<ChartCase>(Total);
            int index = 0;

            foreach (string type in Types)
            {
                foreach (string legend in Legends)
                {
                    foreach (int rows in RowCounts)
                    {
                        foreach (bool labels in new[] { false, true })
                        {
                            list.Add(MakeCreate(index, type, legend, rows, labels));
                            index++;
                        }
                    }
                }
            }

            string[] kinds =
            {
                "values", "values", "to-bar", "to-line", "to-pie2d",
                "to-pie3d", "to-combo-y2", "to-combo-y", "to-twocol", "to-bar"
            };
            for (int i = 0; i < 150; i++)
            {
                string kind = kinds[i % kinds.Length];
                string legend = Legends[i % Legends.Length];
                int rows = RowCounts[i % RowCounts.Length];
                bool labels = i % 2 == 0;
                list.Add(MakeReplace(index, kind, legend, rows, labels, i));
                index++;
            }

            for (int i = 0; i < list.Count; i++)
            {
                list[i].Index = i + 1;
                list[i].Batch = i / BatchSize + 1;
                list[i].Page = i % BatchSize + 1;
            }

            return list;
        }

        private static ChartCase MakeCreate(int seed, string type, string legend, int rows, bool labels)
        {
            List<string> cats = Cats(rows);
            List<double> vals = Vals(seed, rows, 10);
            var spec = new ChartSeriesSpec
            {
                Type = type == "pie" ? "pie2d" : type,
                Axis = "y",
                Name = "系列A",
                Values = vals
            };
            string html = ChartHtml.BuildChart(
                true,
                type,
                legend,
                cats,
                new IList<double>[] { vals },
                new[] { spec.Type },
                new[] { "y" },
                new[] { spec.Name },
                labels);
            return new ChartCase
            {
                Name = "create-" + type + "-" + legend + "-r" + rows + (labels ? "-lbl" : ""),
                IsReplace = false,
                CreateHtml = html,
                ExpectType = type,
                ExpectRows = rows,
                ExpectSeries = 1,
                FirstCategory = cats[0],
                FirstValue = vals[0],
                LastValue = vals[rows - 1]
            };
        }

        private static ChartCase MakeReplace(int index, string kind, string legend, int rows, bool labels, int seed)
        {
            List<string> baseCats = Cats(4);
            List<double> baseVals = Vals(seed, 4, 3);
            string baseline = ChartHtml.BuildChart(
                true,
                "column",
                "none",
                baseCats,
                new IList<double>[] { baseVals },
                new[] { "column" },
                new[] { "y" },
                new[] { "底图" },
                false);

            List<string> cats = Cats(rows);
            string expectType;
            int expectSeries;
            string lineAxis = null;
            IList<IList<double>> series;
            string[] types;
            string[] axes;
            string[] names;
            List<double> primary;

            switch (kind)
            {
                case "to-bar":
                    expectType = "bar";
                    expectSeries = 1;
                    primary = Vals(seed, rows, 20);
                    series = new IList<double>[] { primary };
                    types = new[] { "bar" };
                    axes = new[] { "y" };
                    names = new[] { "条形" };
                    break;
                case "to-line":
                    expectType = "line";
                    expectSeries = 1;
                    primary = Vals(seed, rows, 8);
                    series = new IList<double>[] { primary };
                    types = new[] { "line" };
                    axes = new[] { "y" };
                    names = new[] { "折线" };
                    break;
                case "to-pie2d":
                    expectType = "pie2d";
                    expectSeries = 1;
                    primary = Vals(seed, rows, 15);
                    series = new IList<double>[] { primary };
                    types = new[] { "pie2d" };
                    axes = new[] { "y" };
                    names = new[] { "份额" };
                    break;
                case "to-pie3d":
                    expectType = "pie3d";
                    expectSeries = 1;
                    primary = Vals(seed, rows, 16);
                    series = new IList<double>[] { primary };
                    types = new[] { "pie3d" };
                    axes = new[] { "y" };
                    names = new[] { "份额" };
                    break;
                case "to-combo-y2":
                    expectType = "column";
                    expectSeries = 2;
                    lineAxis = "y2";
                    primary = Vals(seed, rows, 12);
                    series = new IList<double>[] { primary, LineVals(seed, rows) };
                    types = new[] { "column", "line" };
                    axes = new[] { "y", "y2" };
                    names = new[] { "规模", "增速" };
                    break;
                case "to-combo-y":
                    expectType = "column";
                    expectSeries = 2;
                    lineAxis = "y";
                    primary = Vals(seed, rows, 11);
                    series = new IList<double>[] { primary, LineVals(seed, rows) };
                    types = new[] { "column", "line" };
                    axes = new[] { "y", "y" };
                    names = new[] { "规模", "增速" };
                    break;
                case "to-twocol":
                    expectType = "column";
                    expectSeries = 2;
                    primary = Vals(seed, rows, 9);
                    var second = new List<double>(rows);
                    for (int i = 0; i < rows; i++)
                    {
                        second.Add(System.Math.Round(primary[i] * 0.35, 3));
                    }

                    series = new IList<double>[] { primary, second };
                    types = new[] { "column", "column" };
                    axes = new[] { "y", "y" };
                    names = new[] { "国内", "进口" };
                    break;
                default:
                    expectType = "column";
                    expectSeries = 1;
                    primary = Vals(seed, rows, 30);
                    series = new IList<double>[] { primary };
                    types = new[] { "column" };
                    axes = new[] { "y" };
                    names = new[] { "换数" };
                    break;
            }

            string replaceHtml = ChartHtml.BuildChart(
                false,
                expectType,
                legend,
                cats,
                series,
                types,
                axes,
                names,
                labels);
            return new ChartCase
            {
                Name = "replace-" + kind + "-" + legend + "-r" + rows,
                IsReplace = true,
                CreateHtml = baseline,
                ReplaceHtml = replaceHtml,
                ExpectType = expectType,
                ExpectRows = rows,
                ExpectSeries = expectSeries,
                ExpectLineAxis = lineAxis,
                FirstCategory = cats[0],
                FirstValue = primary[0],
                LastValue = primary[rows - 1]
            };
        }

        private static List<string> Cats(int rows)
        {
            var list = new List<string>(rows);
            for (int i = 0; i < rows; i++)
            {
                list.Add((2026 + i).ToString());
            }

            return list;
        }

        private static List<double> Vals(int seed, int rows, int baseVal)
        {
            var list = new List<double>(rows);
            for (int i = 0; i < rows; i++)
            {
                list.Add(System.Math.Round(baseVal + seed % 17 + i * 1.25, 3));
            }

            return list;
        }

        private static List<double> LineVals(int seed, int rows)
        {
            var list = new List<double>(rows);
            for (int i = 0; i < rows; i++)
            {
                list.Add(System.Math.Round(0.2 + seed % 7 * 0.01 + i * 0.012, 3));
            }

            return list;
        }
    }
}
