using System;
using System.Collections.Generic;
using WordAddIn1;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>HTML 铺表：省略成功、占位失败、重叠/未铺满。</summary>
    internal static class TableHtmlGridTest
    {
        public static void Run(TestRun run)
        {
            try
            {
                TestOmitRowspanColspanSuccess();
                TestPlaceholderAfterColspanFails();
                TestRowspanPlaceholderShiftsAndFails();
                TestHoleFailsNotFull();
                TestOverlapFails();
                TestRowspanOutOfBounds();
                TestExtraCellOnLaterRowFails();
                TestBuildDataOnlyOmitsCovered();
                run.Expect("table-html-grid", true, "ok");
            }
            catch (Exception ex)
            {
                run.Fail("table-html-grid", ex.Message);
            }
        }

        private static void TestOmitRowspanColspanSuccess()
        {
            var rows = new List<IList<TableHtmlVisibleCell>>
            {
                new List<TableHtmlVisibleCell>
                {
                    Cell("地区", 3, 1),
                    Cell("产品", 3, 1),
                    Cell("第一季度", 1, 2),
                    Cell("第二季度", 1, 2),
                    Cell("全年合计", 3, 1),
                },
                new List<TableHtmlVisibleCell>
                {
                    Cell("销售额"),
                    Cell("同比增长(%)"),
                    Cell("销售额"),
                    Cell("同比增长(%)"),
                },
                new List<TableHtmlVisibleCell>
                {
                    Cell("销量"),
                    Cell("同比增长(%)"),
                    Cell("销量"),
                    Cell("同比增长(%)"),
                },
            };

            if (!TableHtmlGrid.TryExpand(rows, out TableHtmlGridResult result) || !result.Ok)
            {
                throw new Exception("omit should succeed: " + result.Error);
            }

            if (result.RowCount != 3 || result.ColCount != 7)
            {
                throw new Exception("omit size expected 3x7, got " + result.RowCount + "x" + result.ColCount);
            }

            if (result.RectangularData[0][0] != "地区" || result.RectangularData[0][2] != "第一季度")
            {
                throw new Exception("omit anchor text mismatch");
            }

            if (result.Merge == null || result.Merge.Count < 4)
            {
                throw new Exception("omit should produce merge rects");
            }
        }

        private static void TestPlaceholderAfterColspanFails()
        {
            var rows = new List<IList<TableHtmlVisibleCell>>
            {
                new List<TableHtmlVisibleCell>
                {
                    Cell("地区", 3, 1),
                    Cell("产品", 3, 1),
                    Cell("第一季度", 1, 2),
                    Cell(""),
                    Cell("第二季度", 1, 2),
                    Cell(""),
                    Cell("全年合计", 3, 1),
                },
            };

            if (TableHtmlGrid.TryExpand(rows, out TableHtmlGridResult result) && result.Ok)
            {
                throw new Exception("placeholder after colspan should fail");
            }

            AssertNoPadHint(result.Error);
            if (result.Error == null || result.Error.IndexOf("重叠", StringComparison.Ordinal) < 0)
            {
                throw new Exception("expected overlap, got: " + result.Error);
            }
        }

        private static void TestRowspanPlaceholderShiftsAndFails()
        {
            var rows = new List<IList<TableHtmlVisibleCell>>
            {
                new List<TableHtmlVisibleCell>
                {
                    Cell("A", 2, 1),
                    Cell("B"),
                    Cell("C"),
                },
                new List<TableHtmlVisibleCell>
                {
                    Cell(""),
                    Cell("d"),
                    Cell("e"),
                },
            };

            if (TableHtmlGrid.TryExpand(rows, out TableHtmlGridResult result) && result.Ok)
            {
                throw new Exception("rowspan placeholder should fail");
            }

            AssertNoPadHint(result.Error);
        }

        private static void TestHoleFailsNotFull()
        {
            var rows = new List<IList<TableHtmlVisibleCell>>
            {
                new List<TableHtmlVisibleCell>
                {
                    Cell("A"),
                    Cell("B"),
                    Cell("C"),
                },
                new List<TableHtmlVisibleCell>
                {
                    Cell("d"),
                },
            };

            if (TableHtmlGrid.TryExpand(rows, out TableHtmlGridResult result) && result.Ok)
            {
                throw new Exception("short row should fail");
            }

            if (result.Error == null || result.Error.IndexOf("未铺满逻辑列", StringComparison.Ordinal) < 0)
            {
                throw new Exception("expected 未铺满, got: " + result.Error);
            }

            AssertNoPadHint(result.Error);
        }

        private static void TestOverlapFails()
        {
            var rows = new List<IList<TableHtmlVisibleCell>>
            {
                new List<TableHtmlVisibleCell>
                {
                    Cell("A"),
                    Cell("B", 2, 1),
                    Cell("C"),
                },
                new List<TableHtmlVisibleCell>
                {
                    Cell("x", 1, 2),
                },
            };

            if (TableHtmlGrid.TryExpand(rows, out TableHtmlGridResult result) && result.Ok)
            {
                throw new Exception("overlapping spans should fail");
            }

            AssertNoPadHint(result.Error);
        }

        private static void TestRowspanOutOfBounds()
        {
            var rows = new List<IList<TableHtmlVisibleCell>>
            {
                new List<TableHtmlVisibleCell>
                {
                    Cell("A", 3, 1),
                    Cell("B"),
                },
            };

            if (TableHtmlGrid.TryExpand(rows, out TableHtmlGridResult result) && result.Ok)
            {
                throw new Exception("rowspan beyond rows should fail");
            }

            if (result.Error == null || result.Error.IndexOf("越界", StringComparison.Ordinal) < 0)
            {
                throw new Exception("expected 越界, got: " + result.Error);
            }
        }

        private static void TestExtraCellOnLaterRowFails()
        {
            var rows = new List<IList<TableHtmlVisibleCell>>
            {
                new List<TableHtmlVisibleCell>
                {
                    Cell("A", 2, 1),
                    Cell("B", 1, 2),
                },
                new List<TableHtmlVisibleCell>
                {
                    Cell("c"),
                    Cell("d"),
                    Cell(""),
                },
            };

            if (TableHtmlGrid.TryExpand(rows, out TableHtmlGridResult result) && result.Ok)
            {
                throw new Exception("extra cell on later row should fail");
            }

            AssertNoPadHint(result.Error);
        }

        private static void TestBuildDataOnlyOmitsCovered()
        {
            var dto = new TableExtractDto
            {
                Rows = 2,
                Cols = 2,
                Merge = new List<List<int>> { new List<int> { 1, 1, 2, 1 } },
                Data = new List<List<string>>
                {
                    new List<string> { "A", "B" },
                    new List<string> { "", "C" },
                },
            };

            string xml = TableFormatXmlBuilder.BuildDataOnly(dto);
            if (xml.IndexOf("rowspan=\"2\"", StringComparison.Ordinal) < 0)
            {
                throw new Exception("BuildDataOnly should emit rowspan: " + xml);
            }

            int cellCount = CountOccurrences(xml, "<Cell");
            if (cellCount != 3)
            {
                throw new Exception("BuildDataOnly should omit covered cell, got " + cellCount + ": " + xml);
            }

            if (xml.IndexOf("<Merge>", StringComparison.Ordinal) < 0)
            {
                throw new Exception("BuildDataOnly should keep Properties/Merge");
            }
        }

        private static int CountOccurrences(string text, string token)
        {
            int count = 0;
            int start = 0;
            while (true)
            {
                int i = text.IndexOf(token, start, StringComparison.Ordinal);
                if (i < 0)
                {
                    return count;
                }

                count++;
                start = i + token.Length;
            }
        }

        private static TableHtmlVisibleCell Cell(string text, int rowSpan = 1, int colSpan = 1)
        {
            return new TableHtmlVisibleCell
            {
                Text = text,
                RowSpan = rowSpan,
                ColSpan = colSpan,
            };
        }

        private static void AssertNoPadHint(string error)
        {
            if (string.IsNullOrEmpty(error))
            {
                throw new Exception("expected error");
            }

            string[] banned =
            {
                "请补空",
                "补空 Cell",
                "缺占位格",
                "补空 td",
                "未补空",
                "占位缺失",
            };
            foreach (string item in banned)
            {
                if (error.IndexOf(item, StringComparison.Ordinal) >= 0)
                {
                    throw new Exception("error still teaches padding: " + error);
                }
            }
        }
    }
}
