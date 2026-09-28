using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using WordAddIn1;
using WordAddIn1.HostPlatform;
using WordAddIn1.OpenFiles;
using WordAddIn1.SpreadsheetHost;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWorkbookRoundtripTest
{
    internal sealed class XlSession
    {
        private readonly Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> _tools;

        public XlSession(bool et)
        {
            Et = et;
            App = et ? "et" : "excel";
            ExpectedKind = et ? "et" : "excel";
            _tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.Ordinal);
            F_OpenDocumentTool.Register(_tools, null);
            F_GetWorkbookContentTool.Register(_tools, null);
            F_ReadExcelRangeTool.Register(_tools, null);
            F_WriteExcelRangeTool.Register(_tools, null);
            F_ApplyExcelStructureTool.Register(_tools, null);
            F_ManageExcelSheetTool.Register(_tools, null);
            F_ApplyExcelFormatTool.Register(_tools, null);
            F_GetExcelFormatTool.Register(_tools, null);
            F_ApplyExcelConditionalFormatTool.Register(_tools, null);
        }

        public bool Et { get; }

        public string App { get; }

        public string ExpectedKind { get; }

        public string NewXlsxPath()
        {
            return Path.Combine(
                Path.GetTempPath(),
                "ew-xl-" + Guid.NewGuid().ToString("N") + ".xlsx");
        }

        public ToolResult Call(string name, Dictionary<string, object> args)
        {
            return _tools[name](args ?? new Dictionary<string, object>()).GetAwaiter().GetResult();
        }

        public ToolResult Open(string path, bool createBlank, string app = null)
        {
            return Call("F_open_document", new Dictionary<string, object>
            {
                ["path"] = path,
                ["create_blank"] = createBlank,
                ["app"] = app ?? App
            });
        }

        public ToolResult Overview(string channelId)
        {
            var args = new Dictionary<string, object>();
            if (!string.IsNullOrEmpty(channelId))
            {
                args["channel_id"] = channelId;
            }

            return Call("F_get_workbook_content", args);
        }

        public ToolResult Read(
            string channelId,
            string sheet,
            string range = null,
            string content = null,
            string exportCsv = null)
        {
            var args = new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["sheet"] = sheet
            };
            if (range != null)
            {
                args["range"] = range;
            }

            if (content != null)
            {
                args["content"] = content;
            }

            if (exportCsv != null)
            {
                args["export_csv"] = exportCsv;
            }

            return Call("F_read_excel_range", args);
        }

        public ToolResult Write(
            string channelId,
            string sheet,
            string range,
            string csvPath = null,
            bool clear = false)
        {
            var args = new Dictionary<string, object>
            {
                ["sheet"] = sheet,
                ["range"] = range
            };
            if (!string.IsNullOrEmpty(channelId))
            {
                args["channel_id"] = channelId;
            }

            if (clear)
            {
                args["clear"] = true;
            }
            else if (csvPath != null)
            {
                args["csv_filename"] = csvPath;
            }

            return Call("F_write_excel_range", args);
        }

        public ToolResult Structure(
            string channelId,
            string sheet,
            string action,
            string range,
            int? count = null)
        {
            var args = new Dictionary<string, object>
            {
                ["action"] = action,
                ["range"] = range
            };
            if (!string.IsNullOrEmpty(channelId))
            {
                args["channel_id"] = channelId;
            }

            if (sheet != null)
            {
                args["sheet"] = sheet;
            }

            if (count.HasValue)
            {
                args["count"] = count.Value;
            }

            return Call("F_apply_excel_structure", args);
        }

        public ToolResult Manage(
            string channelId,
            string action,
            string name = null,
            string sheet = null,
            string before = null,
            bool? confirm = null)
        {
            var args = new Dictionary<string, object>();
            if (!string.IsNullOrEmpty(channelId))
            {
                args["channel_id"] = channelId;
            }

            if (action != null)
            {
                args["action"] = action;
            }

            if (name != null)
            {
                args["name"] = name;
            }

            if (sheet != null)
            {
                args["sheet"] = sheet;
            }

            if (before != null)
            {
                args["before"] = before;
            }

            if (confirm.HasValue)
            {
                args["confirm"] = confirm.Value;
            }

            return Call("F_manage_excel_sheet", args);
        }

        public ToolResult ApplyFormat(
            string channelId,
            string sheet,
            string range,
            Dictionary<string, object> format)
        {
            var args = new Dictionary<string, object>
            {
                ["range"] = range,
                ["format"] = format ?? new Dictionary<string, object>()
            };
            if (!string.IsNullOrEmpty(channelId))
            {
                args["channel_id"] = channelId;
            }

            if (sheet != null)
            {
                args["sheet"] = sheet;
            }

            return Call("F_apply_excel_format", args);
        }

        public ToolResult GetFormat(string channelId, string sheet, string range)
        {
            var args = new Dictionary<string, object>
            {
                ["range"] = range
            };
            if (!string.IsNullOrEmpty(channelId))
            {
                args["channel_id"] = channelId;
            }

            if (sheet != null)
            {
                args["sheet"] = sheet;
            }

            return Call("F_get_excel_format", args);
        }

        public ToolResult CondFormat(
            string channelId,
            string sheet,
            string range,
            string action,
            Dictionary<string, object> rule = null)
        {
            var args = new Dictionary<string, object>();
            if (!string.IsNullOrEmpty(channelId))
            {
                args["channel_id"] = channelId;
            }

            if (sheet != null)
            {
                args["sheet"] = sheet;
            }

            if (range != null)
            {
                args["range"] = range;
            }

            if (action != null)
            {
                args["action"] = action;
            }

            if (rule != null)
            {
                args["rule"] = rule;
            }

            return Call("F_apply_excel_conditional_format", args);
        }

        public string NewCsvPath()
        {
            return Path.Combine(
                Path.GetTempPath(),
                "ew-xlw-" + Guid.NewGuid().ToString("N") + ".csv");
        }

        public string WriteCsv(params string[] lines)
        {
            string path = NewCsvPath();
            File.WriteAllText(
                path,
                string.Join("\r\n", lines ?? new string[0]),
                new System.Text.UTF8Encoding(true));
            return path;
        }

        public bool SeedCell(IOperationChannel ch, string sheetName, string a1, object value)
        {
            return WithNamedSheet(ch, sheetName, ws => SetCell(ws, a1, value));
        }

        public bool SeedNumberFormat(IOperationChannel ch, string sheetName, string a1, string format)
        {
            return WithNamedSheet(ch, sheetName, ws =>
            {
                object range = GetRange(ws, a1);
                if (range is Excel.Range excel)
                {
                    excel.NumberFormat = format;
                    return;
                }

                EtCom.TrySetProperty(range, "NumberFormat", format);
            });
        }

        public string ReadNumberFormat(IOperationChannel ch, string sheetName, string a1)
        {
            string format = null;
            WithNamedSheet(ch, sheetName, ws =>
            {
                object range = GetRange(ws, a1);
                if (range is Excel.Range excel)
                {
                    format = Convert.ToString(excel.NumberFormat ?? "");
                    return;
                }

                format = Convert.ToString(EtCom.GetProperty(range, "NumberFormat") ?? "");
            });
            return format ?? "";
        }

        public bool CellHasFormula(IOperationChannel ch, string sheetName, string a1)
        {
            bool has = false;
            WithNamedSheet(ch, sheetName, ws =>
            {
                object range = GetRange(ws, a1);
                if (range is Excel.Range excel)
                {
                    has = excel.HasFormula;
                    return;
                }

                try
                {
                    object raw = EtCom.GetProperty(range, "HasFormula");
                    has = raw != null && Convert.ToBoolean(raw);
                }
                catch
                {
                }
            });
            return has;
        }

        public object ReadValue2(IOperationChannel ch, string sheetName, string a1)
        {
            object value = null;
            WithNamedSheet(ch, sheetName, ws =>
            {
                object range = GetRange(ws, a1);
                if (range is Excel.Range excel)
                {
                    value = excel.Value2;
                    return;
                }

                value = EtCom.GetProperty(range, "Value2");
            });
            return value;
        }

        public bool TryGetChannel(string channelId, out IOperationChannel channel)
        {
            return ChannelRegistry.TryGet(channelId, out channel);
        }

        public void CloseBook(string channelId)
        {
            if (string.IsNullOrEmpty(channelId) || !TryGetChannel(channelId, out IOperationChannel ch))
            {
                return;
            }

            try
            {
                if (ch is ExcelChannel excel && excel.TryGetLiveWorkbook(out Excel.Workbook book))
                {
                    book.Close(false);
                }
                else if (ch is EtChannel et && et.TryGetLiveWorkbook(out object etBook))
                {
                    EtCom.InvokeFlex(etBook, "Close", false);
                }
            }
            catch
            {
            }
        }

        public string FirstSheetName(IOperationChannel ch)
        {
            if (ch is ExcelChannel excel && excel.TryGetLiveWorkbook(out Excel.Workbook book))
            {
                return ((Excel.Worksheet)book.Worksheets[1]).Name;
            }

            if (ch is EtChannel et && et.TryGetLiveWorkbook(out object etBook))
            {
                object sheets = EtCom.GetProperty(etBook, "Worksheets");
                object ws = EtCom.GetIndexed(sheets, 1);
                return Convert.ToString(EtCom.GetProperty(ws, "Name") ?? "");
            }

            return "";
        }

        public bool SeedUsed(IOperationChannel ch, string sheetName)
        {
            return WithSheet(ch, sheetName, ws =>
            {
                RenameSheet(ws, sheetName);
                SetCell(ws, "A1", "地区");
                SetCell(ws, "B1", "销量");
                SetCell(ws, "A2", "东");
                SetCell(ws, "B2", 10);
            });
        }

        public bool SeedFormula(IOperationChannel ch, string sheetName)
        {
            return WithSheet(ch, sheetName, ws =>
            {
                RenameSheet(ws, sheetName);
                SetCell(ws, "A2", 10);
                SetFormula(ws, "B2", "=A2+1");
            });
        }

        public bool SeedMerge(IOperationChannel ch, string sheetName)
        {
            return WithSheet(ch, sheetName, ws =>
            {
                RenameSheet(ws, sheetName);
                object range = GetRange(ws, "A1:C1");
                SetCell(ws, "A1", "标题");
                try
                {
                    if (range is Excel.Range ppt)
                    {
                        ppt.Merge();
                    }
                    else
                    {
                        EtCom.InvokeFlex(range, "Merge");
                    }
                }
                catch
                {
                    return;
                }
            });
        }

        public bool SeedTwoSheets(IOperationChannel ch, string a, string b)
        {
            if (!WithSheet(ch, null, ws =>
                {
                    RenameSheet(ws, a);
                    SetCell(ws, "A1", "利润");
                }))
            {
                return false;
            }

            return AddWorksheet(ch, b, ws => SetCell(ws, "A1", "底稿"));
        }

        public bool SeedHidden(IOperationChannel ch, string vis, string hid)
        {
            if (!SeedTwoSheets(ch, vis, hid))
            {
                return false;
            }

            return WithNamedSheet(ch, hid, ws =>
            {
                try
                {
                    if (ws is Excel.Worksheet excelWs)
                    {
                        excelWs.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                    }
                    else
                    {
                        EtCom.TrySetProperty(ws, "Visible", 0);
                    }
                }
                catch
                {
                }

                SetCell(ws, "A1", "隐");
            });
        }

        public bool SeedWide(IOperationChannel ch, string sheetName)
        {
            return WithSheet(ch, sheetName, ws =>
            {
                RenameSheet(ws, sheetName);
                for (int r = 1; r <= 8; r++)
                {
                    SetCell(ws, "A" + r, "R" + r);
                }

                for (int c = 1; c <= 20; c++)
                {
                    SetCell(ws, A1Address.ColumnLetter(c) + "1", "C" + c);
                }
            });
        }

        public bool SeedChartSheet(IOperationChannel ch, out string chartName)
        {
            chartName = null;
            try
            {
                if (ch is ExcelChannel excel && excel.TryGetLiveWorkbook(out Excel.Workbook book))
                {
                    Excel.Chart chart = (Excel.Chart)book.Charts.Add();
                    chartName = chart.Name;
                    return !string.IsNullOrEmpty(chartName);
                }

                if (ch is EtChannel et && et.TryGetLiveWorkbook(out object etBook))
                {
                    object charts = EtCom.GetProperty(etBook, "Charts");
                    object chart = EtCom.InvokeFlex(charts, "Add");
                    chartName = Convert.ToString(EtCom.GetProperty(chart, "Name") ?? "");
                    return !string.IsNullOrEmpty(chartName);
                }
            }
            catch
            {
            }

            return false;
        }

        public IOperationChannel UnsavedBook()
        {
            if (Et)
            {
                if (!EtCom.TryCreateApplication(out object app, out _))
                {
                    return null;
                }

                object books = EtCom.GetProperty(app, "Workbooks");
                object book = EtCom.InvokeFlex(books, "Add");
                return ChannelRegistry.CreateOrGetEt(book, "");
            }

            if (!ExcelApplicationResolver.TryResolveTyped(out Excel.Application appEx, out _, true))
            {
                return null;
            }

            Excel.Workbook bookEx = appEx.Workbooks.Add();
            return ChannelRegistry.CreateOrGetExcel(bookEx, "");
        }

        private bool AddWorksheet(IOperationChannel ch, string name, Action<object> seed)
        {
            try
            {
                if (ch is ExcelChannel excel && excel.TryGetLiveWorkbook(out Excel.Workbook book))
                {
                    Excel.Worksheet ws = (Excel.Worksheet)book.Worksheets.Add();
                    ws.Name = name;
                    seed(ws);
                    return true;
                }

                if (ch is EtChannel et && et.TryGetLiveWorkbook(out object etBook))
                {
                    object sheets = EtCom.GetProperty(etBook, "Worksheets");
                    object ws = EtCom.InvokeFlex(sheets, "Add");
                    RenameSheet(ws, name);
                    seed(ws);
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private bool WithSheet(IOperationChannel ch, string renameTo, Action<object> act)
        {
            return WithFirstSheet(ch, ws =>
            {
                if (!string.IsNullOrEmpty(renameTo))
                {
                    RenameSheet(ws, renameTo);
                }

                act(ws);
            });
        }

        private bool WithNamedSheet(IOperationChannel ch, string name, Action<object> act)
        {
            try
            {
                if (ch is ExcelChannel excel && excel.TryGetLiveWorkbook(out Excel.Workbook book))
                {
                    act((Excel.Worksheet)book.Worksheets[name]);
                    return true;
                }

                if (ch is EtChannel et && et.TryGetLiveWorkbook(out object etBook))
                {
                    foreach (object sheet in EtCom.EnumerateSheets(etBook))
                    {
                        string n = Convert.ToString(EtCom.GetProperty(sheet, "Name") ?? "");
                        if (string.Equals(n, name, StringComparison.Ordinal))
                        {
                            act(sheet);
                            return true;
                        }
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private bool WithFirstSheet(IOperationChannel ch, Action<object> act)
        {
            try
            {
                if (ch is ExcelChannel excel && excel.TryGetLiveWorkbook(out Excel.Workbook book))
                {
                    act(book.Worksheets[1]);
                    return true;
                }

                if (ch is EtChannel et && et.TryGetLiveWorkbook(out object etBook))
                {
                    object sheets = EtCom.GetProperty(etBook, "Worksheets");
                    act(EtCom.GetIndexed(sheets, 1));
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private static void RenameSheet(object ws, string name)
        {
            if (ws is Excel.Worksheet excel)
            {
                excel.Name = name;
                return;
            }

            EtCom.TrySetProperty(ws, "Name", name);
        }

        private static void SetCell(object ws, string a1, object value)
        {
            object range = GetRange(ws, a1);
            if (range is Excel.Range excel)
            {
                excel.Value2 = value;
                return;
            }

            EtCom.TrySetProperty(range, "Value2", value);
        }

        private static void SetFormula(object ws, string a1, string formula)
        {
            object range = GetRange(ws, a1);
            if (range is Excel.Range excel)
            {
                excel.Formula = formula;
                return;
            }

            EtCom.TrySetProperty(range, "Formula", formula);
        }

        private static object GetRange(object ws, string a1)
        {
            if (ws is Excel.Worksheet excel)
            {
                return excel.Range[a1];
            }

            return EtCom.InvokeFlex(ws, "Range", a1);
        }
    }

    internal sealed class FakeWordChannel : IOperationChannel
    {
        public string ChannelId { get { return "word:fake-xl-test"; } }

        public ChannelKind Kind { get { return ChannelKind.Word; } }
    }
}
