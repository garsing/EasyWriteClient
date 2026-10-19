using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using PptChartRoundtripTest;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWorkbookRoundtripTest
{
    /// <summary>
    /// 同一套 PPT 结构 300 + 属性 250。数据改成格子绑定，在已保存的工作簿上建、换、抽。
    /// --case=xlg550 --host=et|excel。--wh-ids= 与 --wh-from-id= 用套件序号 1–550。
    /// </summary>
    internal static class Xl550Suite
    {
        private const string WpsChartStyle = "148";

        public static void Run(TestRun run, bool et, string[] args)
        {
            string onlyIds = Arg(args, "--wh-ids=");
            int fromId = 0;
            int.TryParse(Arg(args, "--wh-from-id="), NumberStyles.Integer, CultureInfo.InvariantCulture, out fromId);
            HashSet<int> idSet = ParseIds(onlyIds);
            List<ChartCase> built = ChartCaseCatalog.Build();
            built.AddRange(ChartAttrCatalog.Build());
            var cases = new List<ChartCase>();
            var suiteIds = new List<int>();
            for (int n = 0; n < built.Count; n++)
            {
                int suiteId = n + 1;
                if (fromId > 0 && suiteId < fromId)
                {
                    continue;
                }

                if (idSet.Count > 0 && !idSet.Contains(suiteId))
                {
                    continue;
                }

                cases.Add(built[n]);
                suiteIds.Add(suiteId);
            }

            run.TotalTarget = cases.Count;
            Console.WriteLine("表格图 550  " + (et ? "ET" : "Excel") + "  条数=" + cases.Count
                + (idSet.Count == 0 ? "" : "  编号=" + onlyIds)
                + (fromId > 0 ? "  从编号 " + fromId + " 起" : ""));

            var session = new XlSession(et);
            string path = session.NewXlsxPath();
            ToolResult opened = session.Open(path, true);
            if (opened == null || !opened.Success)
            {
                run.CaseFail("xlg550-open", opened == null ? "打开空白簿失败" : (opened.Error ?? "打开空白簿失败"));
                return;
            }

            string channel = Str(opened, "channel_id");
            if (!session.TryGetChannel(channel, out IOperationChannel book))
            {
                run.CaseFail("xlg550-open", "没有工作簿通道");
                return;
            }

            string sheet = session.FirstSheetName(book);
            try
            {
                for (int i = 0; i < cases.Count; i++)
                {
                    Console.WriteLine("  #" + suiteIds[i] + " " + (cases[i].Name ?? "?"));
                    RunOne(run, session, book, channel, sheet, et, suiteIds[i], cases[i]);
                    if ((i + 1) % 25 == 0 || i == cases.Count - 1)
                    {
                        Console.WriteLine("  … " + (i + 1) + "/" + cases.Count
                            + " 过=" + run.CasesPassed + " 败=" + run.CasesFailed);
                    }
                }
            }
            finally
            {
                session.CloseBook(channel);
            }
        }

        private static void RunOne(
            TestRun run,
            XlSession session,
            IOperationChannel book,
            string channel,
            string sheet,
            bool et,
            int suiteId,
            ChartCase one)
        {
            string tag = "xlg550-" + suiteId + "-" + (one.Name ?? "?");
            string createHtml = one.CreateHtml;
            string finalHtml = one.IsReplace ? one.ReplaceHtml : one.CreateHtml;
            if (et)
            {
                ApplyEtChartStyle(one, ref createHtml, ref finalHtml);
            }

            if (!WhPptChartXml.TryExport(finalHtml, out string finalWord, out string error)
                || !TryToCellXml(finalWord, sheet, out string finalXml, out error))
            {
                run.CaseFail(tag, error ?? "转格子图失败");
                return;
            }

            string name = "G" + suiteId.ToString(CultureInfo.InvariantCulture);
            try
            {
                if (!WriteGrid(session, book, channel, sheet, finalWord, out error))
                {
                    run.CaseFail(tag, error ?? "写格子失败");
                    return;
                }

                if (one.IsReplace)
                {
                    if (!WhPptChartXml.TryExport(createHtml, out string baseWord, out error)
                        || !TryToCellXml(baseWord, sheet, out string baseXml, out error))
                    {
                        run.CaseFail(tag, error ?? "底图转格子失败");
                        return;
                    }

                    if (!WriteGrid(session, book, channel, sheet, baseWord, out error))
                    {
                        run.CaseFail(tag, error ?? "写底图格子失败");
                        return;
                    }

                    if (!Create(session, channel, sheet, name, baseXml, out error))
                    {
                        run.CaseFail(tag, "底图失败: " + error);
                        return;
                    }

                    if (SameShape(baseXml, finalXml))
                    {
                        if (!WriteGrid(session, book, channel, sheet, finalWord, out error))
                        {
                            run.CaseFail(tag, error ?? "写换数格子失败");
                            return;
                        }

                        ToolResult applied = session.Chart(channel, "apply", xml: finalXml, chartName: name, scope: "all", sheet: sheet);
                        if (applied == null || !applied.Success)
                        {
                            run.CaseFail(tag, "换数失败: " + (applied == null ? "空结果" : (applied.Error ?? "套用失败")));
                            return;
                        }
                    }
                    else
                    {
                        ToolResult deleted = session.Chart(channel, "delete", chartName: name, sheet: sheet);
                        if (deleted == null || !deleted.Success)
                        {
                            run.CaseFail(tag, "删底图失败: " + (deleted == null ? "空结果" : (deleted.Error ?? "删除失败")));
                            return;
                        }

                        if (!WriteGrid(session, book, channel, sheet, finalWord, out error)
                            || !Create(session, channel, sheet, name, finalXml, out error))
                        {
                            run.CaseFail(tag, "换结构再建失败: " + error);
                            return;
                        }
                    }
                }
                else if (!Create(session, channel, sheet, name, finalXml, out error))
                {
                    run.CaseFail(tag, "建图失败: " + error);
                    return;
                }

                ToolResult extracted = session.Chart(channel, "extract", chartName: name, scope: "all", sheet: sheet);
                string xml = extracted == null ? null : Str(extracted, "xml_content");
                if (extracted == null || !extracted.Success || string.IsNullOrEmpty(xml))
                {
                    run.CaseFail(tag, "抽图失败: " + (extracted == null ? "空结果" : (extracted.Error ?? "无 xml_content")));
                    return;
                }

                string withRows = AttachRows(book, xml);
                string mismatch = WhPptChartXml.Match(one, withRows);
                if (mismatch != null)
                {
                    run.CaseFail(tag, mismatch);
                    return;
                }

                run.CaseOk(tag);
            }
            finally
            {
                try
                {
                    session.Chart(channel, "delete", chartName: name, sheet: sheet);
                }
                catch (Exception)
                {
                }
            }
        }

        private static bool Create(XlSession session, string channel, string sheet, string name, string xml, out string error)
        {
            error = null;
            ToolResult created = session.Chart(
                channel,
                "create",
                Dest(sheet, "AN2:AU16"),
                xml,
                name,
                "all",
                sheet);
            if (created == null || !created.Success)
            {
                error = created == null ? "空结果" : (created.Error ?? "建图失败");
                return false;
            }

            return true;
        }

        private static bool WriteGrid(XlSession session, IOperationChannel book, string channel, string sheet, string wordXml, out string error)
        {
            error = null;
            if (!TryReadGrid(wordXml, out List<string> names, out List<string[]> rows, out error))
            {
                return false;
            }

            session.Write(channel, sheet, "A1:AZ80", clear: true);
            for (int c = 0; c < names.Count; c++)
            {
                session.SeedCell(book, sheet, Cell(c + 1, 1), names[c]);
            }

            for (int r = 0; r < rows.Count; r++)
            {
                string[] line = rows[r];
                for (int c = 0; c < line.Length && c < names.Count; c++)
                {
                    if (!session.SeedCell(book, sheet, Cell(c + 1, r + 2), line[c]))
                    {
                        error = "写格子失败 " + Cell(c + 1, r + 2);
                        return false;
                    }
                }
            }

            string probe = ReadCell(book, sheet, "A2");
            if (string.IsNullOrEmpty(probe))
            {
                error = "A2 写入后读回是空，首格应是 " + (rows.Count > 0 && rows[0].Length > 0 ? rows[0][0] : "");
                return false;
            }

            return true;
        }

        private static bool TryToCellXml(string wordXml, string sheet, out string excelXml, out string error)
        {
            excelXml = null;
            if (!TryReadGrid(wordXml, out List<string> names, out List<string[]> rows, out error))
            {
                return false;
            }

            XDocument doc;
            try
            {
                doc = XDocument.Parse(wordXml);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }

            XElement table = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "DataTable");
            if (table == null)
            {
                error = "没有 DataTable";
                return false;
            }

            XElement header = table.Elements().FirstOrDefault(e => e.Name.LocalName == "Header");
            List<XElement> cols = header == null
                ? new List<XElement>()
                : header.Elements().Where(e => e.Name.LocalName == "Column").ToList();
            if (cols.Count != names.Count || rows.Count < 1)
            {
                error = "格子列或行不够";
                return false;
            }

            int last = rows.Count + 1;
            for (int i = 0; i < cols.Count; i++)
            {
                cols[i].SetAttributeValue("sheet", sheet);
                cols[i].SetAttributeValue("range", Cell(i + 1, 2) + ":" + Cell(i + 1, last));
            }

            XElement oldRows = table.Elements().FirstOrDefault(e => e.Name.LocalName == "Rows");
            if (oldRows != null)
            {
                oldRows.Remove();
            }

            excelXml = doc.ToString(SaveOptions.DisableFormatting);
            return true;
        }

        private static bool TryReadGrid(string wordXml, out List<string> names, out List<string[]> rows, out string error)
        {
            names = new List<string>();
            rows = new List<string[]>();
            error = null;
            XDocument doc;
            try
            {
                doc = XDocument.Parse(wordXml);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }

            XElement header = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "Header");
            if (header != null)
            {
                foreach (XElement col in header.Elements().Where(e => e.Name.LocalName == "Column"))
                {
                    names.Add((col.Value ?? "").Trim());
                }
            }

            foreach (XElement row in doc.Descendants().Where(e => e.Name.LocalName == "Row"))
            {
                rows.Add((row.Value ?? "").Split(','));
            }

            if (names.Count < 2 || rows.Count < 1)
            {
                error = "图数据不足";
                return false;
            }

            return true;
        }

        private static string AttachRows(IOperationChannel book, string xml)
        {
            XDocument doc;
            try
            {
                doc = XDocument.Parse(xml);
            }
            catch (Exception)
            {
                return xml;
            }

            XElement chart = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "Chart");
            XElement table = chart == null ? null : chart.Elements().FirstOrDefault(e => e.Name.LocalName == "DataTable");
            XElement header = table == null ? null : table.Elements().FirstOrDefault(e => e.Name.LocalName == "Header");
            if (header == null)
            {
                return xml;
            }

            List<XElement> cols = header.Elements().Where(e => e.Name.LocalName == "Column").ToList();
            if (cols.Count == 0)
            {
                return xml;
            }

            int height = RangeHeight(Attr(cols[0], "range"));
            if (height < 1)
            {
                return xml;
            }

            var rows = new XElement("Rows");
            for (int r = 0; r < height; r++)
            {
                var cells = new List<string>();
                for (int c = 0; c < cols.Count; c++)
                {
                    cells.Add(ReadCell(book, Attr(cols[c], "sheet"), Offset(Attr(cols[c], "range"), r)));
                }

                rows.Add(new XElement("Row", string.Join(",", cells)));
            }

            table.Add(rows);
            return doc.ToString(SaveOptions.DisableFormatting);
        }

        private static object EtValue(object range)
        {
            foreach (string name in new[] { "Value2", "Value", "Text" })
            {
                try
                {
                    object value = EtCom.GetProperty(range, name);
                    if (value != null && !(value is string s && s.Length == 0))
                    {
                        return value;
                    }
                }
                catch (Exception)
                {
                }
            }

            return null;
        }

        private static string ReadCell(IOperationChannel book, string sheet, string a1)
        {
            if (string.IsNullOrEmpty(sheet) || string.IsNullOrEmpty(a1))
            {
                return "";
            }

            try
            {
                if (book is ExcelChannel excel && excel.TryGetLiveWorkbook(out Excel.Workbook wb))
                {
                    Excel.Worksheet ws = FindExcelSheet(wb, sheet);
                    if (ws == null)
                    {
                        return "";
                    }

                    object value = ((Excel.Range)ws.Range[a1]).Value2;
                    return value == null ? "" : Convert.ToString(value, CultureInfo.InvariantCulture);
                }

                if (book is EtChannel et && et.TryGetLiveWorkbook(out object etBook))
                {
                    object ws = FindEtSheet(etBook, sheet);
                    if (ws == null)
                    {
                        return "";
                    }

                    object range = EtCom.InvokeFlex(ws, "Range", a1);
                    object value = EtValue(range);
                    return value == null ? "" : Convert.ToString(value, CultureInfo.InvariantCulture);
                }
            }
            catch (Exception)
            {
            }

            return "";
        }

        private static Excel.Worksheet FindExcelSheet(Excel.Workbook book, string name)
        {
            foreach (Excel.Worksheet sheet in book.Worksheets)
            {
                if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }

            return null;
        }

        private static object FindEtSheet(object book, string name)
        {
            foreach (object sheet in EtCom.EnumerateSheets(book))
            {
                string sheetName = Convert.ToString(EtCom.GetProperty(sheet, "Name")) ?? "";
                if (string.Equals(sheetName, name, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }

            return null;
        }

        private static int RangeHeight(string range)
        {
            if (string.IsNullOrEmpty(range))
            {
                return 0;
            }

            string[] parts = range.Split(':');
            if (parts.Length != 2)
            {
                return 1;
            }

            int r1 = RowOf(parts[0]);
            int r2 = RowOf(parts[1]);
            return r2 >= r1 ? r2 - r1 + 1 : 0;
        }

        private static string Offset(string range, int rowOffset)
        {
            if (string.IsNullOrEmpty(range))
            {
                return "";
            }

            string start = range.Split(':')[0];
            int row = RowOf(start);
            string col = "";
            for (int i = 0; i < start.Length; i++)
            {
                if (char.IsLetter(start[i]))
                {
                    col += start[i];
                }
            }

            return col + (row + rowOffset).ToString(CultureInfo.InvariantCulture);
        }

        private static int RowOf(string a1)
        {
            int row = 0;
            for (int i = 0; i < a1.Length; i++)
            {
                if (char.IsDigit(a1[i]))
                {
                    row = row * 10 + (a1[i] - '0');
                }
            }

            return row;
        }

        private static string Cell(int col, int row)
        {
            return ColumnName(col) + row.ToString(CultureInfo.InvariantCulture);
        }

        private static string ColumnName(int col)
        {
            string name = "";
            while (col > 0)
            {
                col--;
                name = (char)('A' + (col % 26)) + name;
                col /= 26;
            }

            return name;
        }

        private static bool SameShape(string leftXml, string rightXml)
        {
            try
            {
                XDocument left = XDocument.Parse(leftXml);
                XDocument right = XDocument.Parse(rightXml);
                return string.Equals(TypeOf(left), TypeOf(right), StringComparison.OrdinalIgnoreCase)
                    && ValueCount(left) == ValueCount(right);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string TypeOf(XDocument doc)
        {
            XElement type = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "Type");
            return type == null ? "" : (type.Value ?? "").Trim();
        }

        private static int ValueCount(XDocument doc)
        {
            return doc.Descendants().Count(e =>
                e.Name.LocalName == "Column"
                && e.Attribute("type") != null
                && string.Equals(e.Attribute("type").Value, "value", StringComparison.OrdinalIgnoreCase));
        }

        private static void ApplyEtChartStyle(ChartCase one, ref string createHtml, ref string finalHtml)
        {
            if (one == null || one.Attrs == null)
            {
                return;
            }

            bool styleCase = false;
            for (int i = 0; i < one.Attrs.Count; i++)
            {
                AttrCheck check = one.Attrs[i];
                if (check == null
                    || !string.Equals(check.Key, "ChartStyle", StringComparison.Ordinal)
                    || !string.Equals(check.Expected, "286", StringComparison.Ordinal))
                {
                    continue;
                }

                check.Expected = WpsChartStyle;
                styleCase = true;
            }

            if (!styleCase)
            {
                return;
            }

            createHtml = SwapChartStyle(createHtml);
            finalHtml = SwapChartStyle(finalHtml);
        }

        private static string SwapChartStyle(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return html;
            }

            return html.Replace("data-chart-style=\"286\"", "data-chart-style=\"" + WpsChartStyle + "\"");
        }

        private static Dictionary<string, object> Dest(string sheet, string range)
        {
            return new Dictionary<string, object>
            {
                ["sheet"] = sheet,
                ["range"] = range
            };
        }

        private static string Attr(XElement el, string name)
        {
            if (el == null)
            {
                return "";
            }

            XAttribute a = el.Attributes().FirstOrDefault(x =>
                string.Equals(x.Name.LocalName, name, StringComparison.OrdinalIgnoreCase));
            return a == null ? "" : (a.Value ?? "");
        }

        private static string Str(ToolResult result, string key)
        {
            var data = result == null ? null : result.Data as Dictionary<string, object>;
            if (data == null || !data.ContainsKey(key) || data[key] == null)
            {
                return "";
            }

            return Convert.ToString(data[key]) ?? "";
        }

        private static string Arg(string[] args, string prefix)
        {
            if (args == null)
            {
                return "";
            }

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i] ?? "";
                if (a.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    return a.Substring(prefix.Length).Trim();
                }
            }

            return "";
        }

        private static HashSet<int> ParseIds(string onlyIds)
        {
            var set = new HashSet<int>();
            if (string.IsNullOrWhiteSpace(onlyIds))
            {
                return set;
            }

            string[] parts = onlyIds.Split(',');
            for (int i = 0; i < parts.Length; i++)
            {
                if (int.TryParse(parts[i].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int id))
                {
                    set.Add(id);
                }
            }

            return set;
        }
    }
}
