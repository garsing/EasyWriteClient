using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using WordAddIn1;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    internal sealed class WdSession
    {
        public WdSession(string host)
        {
            Host = string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase) ? "wps" : "word";
            App = Host;
            ExpectedKind = Host;
        }

        public string Host { get; }

        public string App { get; }

        public string ExpectedKind { get; }

        public bool IsWps
        {
            get { return Host == "wps"; }
        }

        /// <summary>关稿后顺手收 Word AddChart 留下的无窗口空 Excel。图表波打开，不 Quit 有窗的用户 Excel。</summary>
        public bool DismissOrphanExcelOnClose { get; set; }

        public string NewDocxPath()
        {
            return Path.Combine(
                Path.GetTempPath(),
                "ew-wd-" + Guid.NewGuid().ToString("N") + ".docx");
        }

        public static string DummyPath(string ext)
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "ew-wd-dummy-" + Guid.NewGuid().ToString("N") + ext);
            File.WriteAllBytes(path, new byte[] { 0 });
            return path;
        }

        public ToolResult Open(string path, bool? createBlank = null, string app = null)
        {
            var args = new Dictionary<string, object>();
            if (path != null)
            {
                args["path"] = path;
            }

            if (createBlank.HasValue)
            {
                args["create_blank"] = createBlank.Value;
            }

            if (app != null)
            {
                args["app"] = app;
            }

            return DesktopToolPipe.Call("F_open_document", args);
        }

        public ToolResult PageInfo(string channelId)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return DesktopToolPipe.Call("F_get_document_page_info", args);
        }

        public ToolResult GetContent(string channelId, string codeLevel = null)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            if (codeLevel != null)
            {
                args["code_level"] = codeLevel;
            }

            return DesktopToolPipe.Call("F_get_document_content", args);
        }

        public ToolResult ProcessActions(string channelId, IList<Dictionary<string, object>> actions)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            args["action"] = actions ?? new List<Dictionary<string, object>>();
            return DesktopToolPipe.Call("F_process_document_actions", args);
        }

        public ToolResult ProcessActionsRaw(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_process_document_actions",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult InsertFromFile(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_insert_document_from_file",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult ReplaceFromFile(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_replace_document_from_file",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult GetFormatContext(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_get_format_context",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult ApplyDocumentFormat(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_apply_document_format",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult ApplyParagraphFormat(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_apply_paragraph_format",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult ApplyPageSetup(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_apply_page_setup",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult CreateTableFromXml(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_create_table_from_xml",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult DeleteTable(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_delete_table",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult ReadTableRowValues(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_read_table_row_values",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult ApplyTableRowValues(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_apply_table_row_values",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult ExtractTableFormat(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_extract_table_format",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult ApplyTableFormatFromXml(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_apply_table_format_from_xml",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult OptimizeTableDisplay(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_optimize_table_display",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult CreateChartFromXml(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_create_chart_from_xml",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult ExtractChartXml(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_extract_chart_xml",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult ApplyChartFromXml(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_apply_chart_from_xml",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult DeleteChart(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_delete_chart",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult ListOperations(string channelId)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return DesktopToolPipe.Call("F_list_document_operations", args);
        }

        public ToolResult RestoreCheckpoint(string channelId, string timestamp)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            if (timestamp != null)
            {
                args["timestamp"] = timestamp;
            }

            return DesktopToolPipe.Call("F_restore_document_checkpoint", args);
        }

        public ToolResult ApplyDataProvenance(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_apply_data_provenance",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult ClearDataProvenance(string channelId)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return DesktopToolPipe.Call("F_clear_data_provenance", args);
        }

        public ToolResult GetDataProvenance(string channelId)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return DesktopToolPipe.Call("F_get_data_provenance", args);
        }

        public ToolResult InsertImage(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_insert_image",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult InsertSvgImage(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_insert_svg_image",
                args ?? new Dictionary<string, object>());
        }

        public ToolResult CapturePageImage(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call(
                "F_capture_document_page_image",
                args ?? new Dictionary<string, object>());
        }

        public bool ActivateDoc(string fullPath)
        {
            return IsWps
                ? WithWpsDoc(fullPath, doc =>
                {
                    Invoke(doc, "Activate");
                    return true;
                })
                : WithWordDoc(fullPath, doc =>
                {
                    doc.Activate();
                    return true;
                });
        }

        public ToolResult InsertPageBreakTool(string channelId, Dictionary<string, object> target)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            if (target != null)
            {
                foreach (KeyValuePair<string, object> kv in target)
                {
                    args[kv.Key] = kv.Value;
                }
            }

            return DesktopToolPipe.Call("F_insert_page_break", args);
        }

        public bool MoveSelectionToEnd(string fullPath)
        {
            return IsWps
                ? WithWpsDoc(fullPath, doc =>
                {
                    object app = GetProp(doc, "Application");
                    object sel = GetProp(app, "Selection");
                    object content = GetProp(doc, "Content");
                    int end = Convert.ToInt32(GetProp(content, "End"));
                    Invoke(sel, "SetRange", end, end);
                    return true;
                })
                : WithWordDoc(fullPath, doc =>
                {
                    int end = doc.Content.End;
                    doc.Application.Selection.SetRange(end, end);
                    return true;
                });
        }

        public bool WriteSelectionFact(string fullPath)
        {
            return IsWps
                ? WithWpsDoc(fullPath, doc =>
                {
                    object app = GetProp(doc, "Application");
                    object sel = GetProp(app, "Selection");
                    Console.WriteLine("Selection=" + GetProp(sel, "Start") + "-" + GetProp(sel, "End"));
                    return true;
                })
                : WithWordDoc(fullPath, doc =>
                {
                    Word.Selection sel = doc.Application.Selection;
                    Console.WriteLine("Selection=" + sel.Start + "-" + sel.End);
                    return true;
                });
        }

        public bool WriteTableFacts(string fullPath)
        {
            return IsWps
                ? WithWpsDoc(fullPath, doc =>
                {
                    object tables = GetProp(doc, "Tables");
                    int n = Convert.ToInt32(GetProp(tables, "Count"));
                    Console.WriteLine("COM Tables.Count=" + n);
                    for (int i = 1; i <= n; i++)
                    {
                        object t = Invoke(tables, "Item", i);
                        object rng = GetProp(t, "Range");
                        string title = "";
                        try { title = Convert.ToString(GetProp(t, "Title")); } catch { title = "?"; }
                        string cell = "";
                        try
                        {
                            object c11 = Invoke(t, "Cell", 1, 1);
                            object cr = GetProp(c11, "Range");
                            cell = (Convert.ToString(GetProp(cr, "Text")) ?? "")
                                .Replace("\r", "\\r").Replace("\a", "\\a");
                        }
                        catch (Exception ex)
                        {
                            cell = ex.Message;
                        }

                        Console.WriteLine("  COM[" + i + "] start=" + GetProp(rng, "Start")
                            + " end=" + GetProp(rng, "End")
                            + " title=" + title
                            + " r1=" + cell);
                    }

                    return true;
                })
                : WithWordDoc(fullPath, doc =>
                {
                    Console.WriteLine("COM Tables.Count=" + doc.Tables.Count);
                    for (int i = 1; i <= doc.Tables.Count; i++)
                    {
                        Word.Table t = doc.Tables[i];
                        string title = "";
                        try { title = t.Title ?? ""; } catch { title = "?"; }
                        string cell = "";
                        try
                        {
                            cell = (t.Cell(1, 1).Range.Text ?? "").Replace("\r", "\\r").Replace("\a", "\\a");
                        }
                        catch (Exception ex)
                        {
                            cell = ex.Message;
                        }

                        Console.WriteLine("  COM[" + i + "] start=" + t.Range.Start
                            + " end=" + t.Range.End
                            + " title=" + title
                            + " r1=" + cell);
                    }

                    return true;
                });
        }

        public bool OnWordDoc(string fullPath, Func<Word.Document, bool> action)
        {
            if (IsWps || action == null)
            {
                return false;
            }

            return WithWordDoc(fullPath, action);
        }

        public string GetLiveWordOpenXml(string fullPath)
        {
            string xml = null;
            if (IsWps)
            {
                WithWpsDoc(fullPath, doc =>
                {
                    object content = GetProp(doc, "Content");
                    object raw = GetProp(content, "WordOpenXML");
                    xml = raw as string ?? (raw != null ? raw.ToString() : "");
                    return true;
                });
            }
            else
            {
                WithWordDoc(fullPath, doc =>
                {
                    xml = doc.Content.WordOpenXML;
                    return true;
                });
            }

            return xml ?? "";
        }

        public static bool XmlHasPageBreak(string xml)
        {
            if (string.IsNullOrEmpty(xml))
            {
                return false;
            }

            return xml.IndexOf("type=\"page\"", StringComparison.OrdinalIgnoreCase) >= 0
                || xml.IndexOf("type='page'", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool XmlHasParagraphSectPr(string xml)
        {
            if (string.IsNullOrEmpty(xml))
            {
                return false;
            }

            return Regex.IsMatch(
                xml,
                @"<w:pPr[\s\S]*?<w:sectPr[\s\S]*?</w:pPr>",
                RegexOptions.IgnoreCase);
        }

        public bool TypePlain(string fullPath, string text)
        {
            return IsWps
                ? WithWpsDoc(fullPath, doc =>
                {
                    object content = GetProp(doc, "Content");
                    SetProp(content, "Text", text ?? "");
                    return true;
                })
                : WithWordDoc(fullPath, doc =>
                {
                    doc.Content.Text = text ?? "";
                    return true;
                });
        }

        public bool SetParagraphChar(string fullPath, int paraIndex, string fontName, float size, bool bold)
        {
            return IsWps
                ? WithWpsDoc(fullPath, doc =>
                {
                    object range = ParagraphRangeLate(doc, paraIndex);
                    if (range == null)
                    {
                        return false;
                    }

                    ApplyFontLate(range, fontName, size, bold);
                    return true;
                })
                : WithWordDoc(fullPath, doc =>
                {
                    if (paraIndex < 1 || paraIndex > doc.Paragraphs.Count)
                    {
                        return false;
                    }

                    Word.Range range = doc.Paragraphs[paraIndex].Range;
                    if (!string.IsNullOrEmpty(fontName))
                    {
                        range.Font.NameFarEast = fontName;
                        range.Font.Name = fontName;
                    }

                    if (size > 0)
                    {
                        range.Font.Size = size;
                    }

                    range.Font.Bold = bold ? 1 : 0;
                    return true;
                });
        }

        public bool SetParagraphAlign(string fullPath, int paraIndex, int alignment)
        {
            return IsWps
                ? WithWpsDoc(fullPath, doc =>
                {
                    object range = ParagraphRangeLate(doc, paraIndex);
                    if (range == null)
                    {
                        return false;
                    }

                    object pf = GetProp(range, "ParagraphFormat");
                    SetProp(pf, "Alignment", alignment);
                    return true;
                })
                : WithWordDoc(fullPath, doc =>
                {
                    if (paraIndex < 1 || paraIndex > doc.Paragraphs.Count)
                    {
                        return false;
                    }

                    doc.Paragraphs[paraIndex].Range.ParagraphFormat.Alignment =
                        (Word.WdParagraphAlignment)alignment;
                    return true;
                });
        }

        public bool TryReadLeftMargin(string fullPath, out double points)
        {
            double value = 0;
            bool ok = IsWps
                ? WithWpsDoc(fullPath, doc =>
                {
                    object ps = GetProp(doc, "PageSetup");
                    value = Convert.ToDouble(GetProp(ps, "LeftMargin"));
                    return true;
                })
                : WithWordDoc(fullPath, doc =>
                {
                    value = doc.PageSetup.LeftMargin;
                    return true;
                });
            points = value;
            return ok;
        }

        public bool InsertTable(string fullPath, int rows, int cols, string[] cells)
        {
            return IsWps
                ? WithWpsDoc(fullPath, doc => InsertTableLate(doc, rows, cols, cells))
                : WithWordDoc(fullPath, doc =>
                {
                    int end = doc.Content.End;
                    Word.Range range = doc.Range(end - 1, end - 1);
                    Word.Table table = doc.Tables.Add(range, rows, cols);
                    int n = cells == null ? 0 : cells.Length;
                    int k = 0;
                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            string t = k < n ? (cells[k] ?? "") : "";
                            k++;
                            table.Cell(r, c).Range.Text = t;
                        }
                    }

                    return true;
                });
        }

        public bool InsertPicture(string fullPath)
        {
            string png = WriteOnePxPng();
            return IsWps
                ? WithWpsDoc(fullPath, doc =>
                {
                    object shapes = GetProp(doc, "InlineShapes");
                    Invoke(shapes, "AddPicture", png);
                    return true;
                })
                : WithWordDoc(fullPath, doc =>
                {
                    doc.InlineShapes.AddPicture(png);
                    return true;
                });
        }

        public bool InsertChart(string fullPath)
        {
            return IsWps
                ? WithWpsDoc(fullPath, doc =>
                {
                    try
                    {
                        object shapes = GetProp(doc, "InlineShapes");
                        Invoke(shapes, "AddChart", 51);
                        return true;
                    }
                    catch
                    {
                    }

                    try
                    {
                        object shapes = GetProp(doc, "Shapes");
                        Invoke(shapes, "AddChart", 51);
                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                })
                : WithWordDoc(fullPath, doc =>
                {
                    doc.InlineShapes.AddChart(Microsoft.Office.Core.XlChartType.xlColumnClustered);
                    return true;
                });
        }

        /// <summary>用本机 Word 做好含 Chart 的 docx 夹具，不 Quit Word。</summary>
        public string CreateChartFixtureDocx(int charts = 1)
        {
            return CreateWordDocxFixture(doc =>
            {
                int n = charts < 1 ? 1 : charts;
                for (int i = 0; i < n; i++)
                {
                    Word.InlineShape shape = doc.InlineShapes.AddChart(
                        Microsoft.Office.Core.XlChartType.xlColumnClustered);
                    if (!StampChartSeriesName(shape, "加固系列" + (i + 1)))
                    {
                        throw new InvalidOperationException("StampChart 失败");
                    }
                }
            });
        }

        public string CreateTwoTableFixtureDocx()
        {
            return CreateWordDocxFixture(doc =>
            {
                doc.Content.Text = "表前间隔。\r表间隔。\r";
                Word.Table first = doc.Tables.Add(doc.Paragraphs[1].Range, 2, 2);
                FillWordTable(first, new[] { "甲", "乙", "丙", "丁" });
                Word.Table second = doc.Tables.Add(doc.Paragraphs[doc.Paragraphs.Count].Range, 2, 2);
                FillWordTable(second, new[] { "戊", "己", "庚", "辛" });
                if (doc.Tables.Count < 2)
                {
                    throw new InvalidOperationException("双表夹具只得到 " + doc.Tables.Count + " 张表");
                }
            });
        }

        public string CreateMixFixtureDocx()
        {
            return CreateWordDocxFixture(doc =>
            {
                doc.Content.Text = "易写混合甲句。";
                int end = doc.Content.End;
                Word.Range range = doc.Range(end - 1, end - 1);
                Word.Table table = doc.Tables.Add(range, 2, 2);
                table.Cell(1, 1).Range.Text = "甲";
                table.Cell(1, 2).Range.Text = "乙";
                table.Cell(2, 1).Range.Text = "丙";
                table.Cell(2, 2).Range.Text = "丁";
                doc.InlineShapes.AddPicture(WriteOnePxPng());
                Word.InlineShape shape = doc.InlineShapes.AddChart(
                    Microsoft.Office.Core.XlChartType.xlColumnClustered);
                if (!StampChartSeriesName(shape, "混合图"))
                {
                    throw new InvalidOperationException("StampChart 失败");
                }
            });
        }

        public bool InsertCharts(string fullPath, int count)
        {
            int n = count < 1 ? 1 : count;
            return IsWps
                ? false
                : WithWordDoc(fullPath, doc =>
                {
                    for (int i = 0; i < n; i++)
                    {
                        Word.InlineShape shape = doc.InlineShapes.AddChart(
                            Microsoft.Office.Core.XlChartType.xlColumnClustered);
                        if (!StampChartSeriesName(shape, "加固系列" + (i + 1)))
                        {
                            return false;
                        }
                    }

                    return true;
                });
        }

        private string CreateWordDocxFixture(Action<Word.Document> plant)
        {
            Word.Application app = null;
            bool createdApp = false;
            try
            {
                try
                {
                    app = (Word.Application)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    app = new Word.Application();
                    app.Visible = true;
                    createdApp = true;
                }

                Word.Document doc = app.Documents.Add();
                plant(doc);
                string path = NewDocxPath();
                object fileName = path;
                doc.SaveAs2(ref fileName);
                object save = false;
                doc.Close(ref save);
                return path;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (createdApp)
                {
                    // 本波自己拉起的 Word 只关稿，不 Quit。
                }
            }
        }

        private static bool StampChartSeriesName(Word.InlineShape shape, string name)
        {
            try
            {
                dynamic chart = shape.Chart;
                dynamic series = chart.SeriesCollection(1);
                series.Name = name;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool IsShortId(string id)
        {
            return !string.IsNullOrEmpty(id)
                && Regex.IsMatch(id, @"^w\d+$", RegexOptions.CultureInvariant);
        }

        public static Dictionary<string, object> Data(ToolResult r)
        {
            return r == null ? null : r.Data as Dictionary<string, object>;
        }

        public static string Str(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return "";
            }

            return Convert.ToString(d[key]) ?? "";
        }

        public static bool Bool(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return false;
            }

            object v = d[key];
            if (v is bool b)
            {
                return b;
            }

            return string.Equals(Convert.ToString(v), "true", StringComparison.OrdinalIgnoreCase);
        }

        public static int IntOf(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToInt32(d[key]);
            }
            catch
            {
                return 0;
            }
        }

        public static bool HasKey(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            return d != null && d.ContainsKey(key);
        }

        public bool InsertPageBreak(string fullPath)
        {
            return IsWps ? InsertPageBreakWps(fullPath) : InsertPageBreakWord(fullPath);
        }

        /// <summary>
        /// 封面句 + 分页 + 概述句 + 下一页分节 + 正文句。供读文后再 Find。
        /// </summary>
        public bool WriteLayoutBreakFixture(string fullPath)
        {
            return IsWps
                ? WithWpsDoc(fullPath, WriteLayoutBreakFixtureLate)
                : WithWordDoc(fullPath, WriteLayoutBreakFixtureWord);
        }

        /// <summary>
        /// 用现网 Find 路径搜存文。miss 为转义后的搜不到的存文。
        /// </summary>
        public bool FindAllStoredTexts(string fullPath, IList<string> storedTexts, out string miss)
        {
            miss = null;
            if (storedTexts == null || storedTexts.Count == 0)
            {
                miss = "(无存文)";
                return false;
            }

            string failed = null;
            bool ok = IsWps
                ? WithWpsDoc(fullPath, doc => FindAllStoredTextsLate(doc, storedTexts, out failed))
                : WithWordDoc(fullPath, doc =>
                {
                    foreach (string stored in storedTexts)
                    {
                        if (string.IsNullOrEmpty(stored))
                        {
                            continue;
                        }

                        Word.Range hit = WordRangeFinder.FindFirstInDocument(doc, stored, "layout-break");
                        if (hit == null)
                        {
                            failed = TextUtils.EscapeNewlinesForDb(stored)
                                + " → "
                                + WordRangeFinder.ConvertNewlinesToWordCodes(stored);
                            return false;
                        }
                    }

                    return true;
                });

            miss = failed;
            return ok;
        }

        public bool CloseDocByPath(string fullPath)
        {
            bool closed = IsWps ? CloseWpsByPath(fullPath) : CloseWordByPath(fullPath);
            TryDeleteTestTemp(fullPath);
            if (DismissOrphanExcelOnClose)
            {
                DismissOrphanExcel();
            }

            return closed;
        }

        /// <summary>开跑前：关残留测试稿、收无窗空 Excel、删 TEMP 里 ew- 测试文件。不 Quit Word/WPS，不碰有窗 Excel。</summary>
        public static void SweepBeforeRun(string host)
        {
            ReportSweep("开跑前", SweepResidues(host));
        }

        /// <summary>跑完再扫一遍，避免本轮关稿失败留下的文件/空 Excel。</summary>
        public static void SweepAfterRun(string host)
        {
            ReportSweep("跑完", SweepResidues(host));
        }

        private static void ReportSweep(string when, int[] counts)
        {
            Console.WriteLine(
                "清残留 " + when
                + "：关测试稿 " + counts[0]
                + "，空 Excel " + counts[1]
                + "，删临时文件 " + counts[2]);
        }

        private static int[] SweepResidues(string host)
        {
            bool wps = string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase);
            int docs = CloseLeftoverTestDocs(wps);
            int excel = DismissOrphanExcel();
            int files = SweepTestTempFiles();
            return new[] { docs, excel, files };
        }

        public static int CloseLeftoverTestDocs(bool wps)
        {
            return wps ? CloseLeftoverWpsTestDocs() : CloseLeftoverWordTestDocs();
        }

        private static int CloseLeftoverWordTestDocs()
        {
            int n = 0;
            try
            {
                var app = (Word.Application)Marshal.GetActiveObject("Word.Application");
                var list = new List<Word.Document>();
                foreach (Word.Document d in app.Documents)
                {
                    list.Add(d);
                }

                foreach (Word.Document d in list)
                {
                    string full = null;
                    string name = null;
                    try
                    {
                        full = d.FullName;
                        name = d.Name;
                    }
                    catch
                    {
                    }

                    if (!IsTestTempName(name) && !IsTestTempPath(full))
                    {
                        continue;
                    }

                    try
                    {
                        object save = false;
                        d.Close(ref save);
                        n++;
                    }
                    catch
                    {
                    }

                    TryDeleteTestTemp(full);
                }
            }
            catch
            {
            }

            return n;
        }

        private static int CloseLeftoverWpsTestDocs()
        {
            int n = 0;
            foreach (object app in EnumerateWpsApps())
            {
                object docs;
                try
                {
                    docs = GetProp(app, "Documents");
                }
                catch
                {
                    continue;
                }

                int count;
                try
                {
                    count = Convert.ToInt32(GetProp(docs, "Count"));
                }
                catch
                {
                    continue;
                }

                var paths = new List<string>();
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        object d = Invoke(docs, "Item", i);
                        string full = Convert.ToString(GetProp(d, "FullName"));
                        string name = Convert.ToString(GetProp(d, "Name"));
                        if (IsTestTempName(name) || IsTestTempPath(full))
                        {
                            paths.Add(full);
                        }
                    }
                    catch
                    {
                    }
                }

                foreach (string full in paths)
                {
                    if (CloseWpsByPath(full))
                    {
                        n++;
                    }

                    TryDeleteTestTemp(full);
                }
            }

            return n;
        }

        /// <summary>
        /// 只收无窗口、且没有用户工作簿的 Excel（Word AddChart 数据源）。
        /// 有可见窗口或磁盘上非测试文件的工作簿一律不动。
        /// </summary>
        public static int DismissOrphanExcel()
        {
            int n = KillWindowlessExcel();
            if (!HasUserExcelWindow() && HasLiveExcelProcess())
            {
                n += QuitEmptyExcelApps();
            }

            return n;
        }

        private static int QuitEmptyExcelApps()
        {
            int quit = 0;
            for (int i = 0; i < 3; i++)
            {
                Excel.Application app;
                try
                {
                    app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    break;
                }

                try
                {
                    if (HasUserWorkbook(app))
                    {
                        break;
                    }

                    var books = new List<Excel.Workbook>();
                    foreach (Excel.Workbook book in app.Workbooks)
                    {
                        books.Add(book);
                    }

                    foreach (Excel.Workbook book in books)
                    {
                        try
                        {
                            book.Close(false);
                        }
                        catch
                        {
                        }
                    }

                    app.DisplayAlerts = false;
                    app.Quit();
                    quit++;
                }
                catch
                {
                    break;
                }
            }

            return quit;
        }

        /// <summary>标题含 xlsx/工作簿/BookN 视为用户 Excel。光秃「Excel」是图表数据窗，不算。</summary>
        private static bool HasLiveExcelProcess()
        {
            try
            {
                foreach (Process p in Process.GetProcessesByName("EXCEL"))
                {
                    try
                    {
                        if (p.WorkingSet64 > 1024 * 1024)
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        p.Dispose();
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool HasUserExcelWindow()
        {
            try
            {
                foreach (Process p in Process.GetProcessesByName("EXCEL"))
                {
                    try
                    {
                        string title = p.MainWindowTitle ?? "";
                        if (LooksLikeUserExcelTitle(title))
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        p.Dispose();
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool LooksLikeUserExcelTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                return false;
            }

            if (title.IndexOf(".xls", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (title.IndexOf("工作簿", StringComparison.Ordinal) >= 0)
            {
                return true;
            }

            return title.IndexOf("Book", StringComparison.OrdinalIgnoreCase) >= 0
                && title.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static int KillWindowlessExcel()
        {
            int n = 0;
            try
            {
                foreach (Process p in Process.GetProcessesByName("EXCEL"))
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(p.MainWindowTitle))
                        {
                            continue;
                        }

                        p.Kill();
                        n++;
                    }
                    catch
                    {
                    }
                    finally
                    {
                        p.Dispose();
                    }
                }
            }
            catch
            {
            }

            return n;
        }

        private static bool HasUserWorkbook(Excel.Application app)
        {
            try
            {
                foreach (Excel.Workbook book in app.Workbooks)
                {
                    try
                    {
                        string full = book.FullName;
                        if (!string.IsNullOrEmpty(full)
                            && File.Exists(full)
                            && !IsTestTempPath(full))
                        {
                            return true;
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        public static int SweepTestTempFiles()
        {
            int n = 0;
            string dir;
            try
            {
                dir = Path.GetTempPath();
            }
            catch
            {
                return 0;
            }

            string[] names;
            try
            {
                names = Directory.GetFiles(dir);
            }
            catch
            {
                return 0;
            }

            foreach (string file in names)
            {
                if (!IsTestTempPath(file))
                {
                    continue;
                }

                if (TryDeleteFile(file))
                {
                    n++;
                }
            }

            return n;
        }

        private static bool IsTestTempName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            return name.StartsWith("ew-wd-", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("ew-wh-", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("ew-c2-", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("ew-wt-", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("ew-wf-", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("ew-wg-", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("ew-wa-", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("ew-wc-", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTestTempPath(string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath))
            {
                return false;
            }

            string name = Path.GetFileName(fullPath);
            if (!IsTestTempName(name))
            {
                return false;
            }

            string temp;
            try
            {
                temp = Path.GetTempPath();
            }
            catch
            {
                return false;
            }

            return string.Equals(
                Normalize(Path.GetDirectoryName(fullPath) ?? ""),
                Normalize(temp.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)),
                StringComparison.OrdinalIgnoreCase);
        }

        private static void TryDeleteTestTemp(string fullPath)
        {
            if (IsTestTempPath(fullPath))
            {
                TryDeleteFile(fullPath);
            }
        }

        private static bool TryDeleteFile(string fullPath)
        {
            for (int i = 0; i < 4; i++)
            {
                try
                {
                    if (!File.Exists(fullPath))
                    {
                        return i > 0;
                    }

                    File.Delete(fullPath);
                    return true;
                }
                catch
                {
                    if (i < 3)
                    {
                        System.Threading.Thread.Sleep(80);
                    }
                }
            }

            return false;
        }

        public static bool CloseExcelByPath(string fullPath)
        {
            try
            {
                var app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                string norm = Normalize(fullPath);
                foreach (Excel.Workbook book in app.Workbooks)
                {
                    try
                    {
                        if (string.Equals(Normalize(book.FullName), norm, StringComparison.OrdinalIgnoreCase))
                        {
                            book.Close(false);
                            TryDeleteTestTemp(fullPath);
                            return true;
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        public static bool ClosePptByPath(string fullPath)
        {
            try
            {
                var app = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                string norm = Normalize(fullPath);
                foreach (PowerPoint.Presentation pres in app.Presentations)
                {
                    try
                    {
                        if (string.Equals(Normalize(pres.FullName), norm, StringComparison.OrdinalIgnoreCase))
                        {
                            pres.Close();
                            TryDeleteTestTemp(fullPath);
                            return true;
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool WriteLayoutBreakFixtureWord(Word.Document doc)
        {
            doc.Content.Text = "封面日期甲。";
            object page = Word.WdBreakType.wdPageBreak;
            doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertBreak(ref page);
            Word.Range afterPage = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
            afterPage.Text = "总概述乙。";
            object sect = Word.WdBreakType.wdSectionBreakNextPage;
            doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertBreak(ref sect);
            Word.Range afterSect = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
            afterSect.Text = "第一章丙。";
            return true;
        }

        private static bool WriteLayoutBreakFixtureLate(object doc)
        {
            object content = GetProp(doc, "Content");
            SetProp(content, "Text", "封面日期甲。");
            InsertBreakLate(doc, 7);
            InsertAfterEndLate(doc, "总概述乙。");
            InsertBreakLate(doc, 2);
            InsertAfterEndLate(doc, "第一章丙。");
            return true;
        }

        private static void InsertAfterEndLate(object doc, string text)
        {
            object content = GetProp(doc, "Content");
            int end = Convert.ToInt32(GetProp(content, "End"));
            int pos = end > 1 ? end - 1 : end;
            object range = Invoke(doc, "Range", pos, pos);
            SetProp(range, "Text", text);
        }

        private static void InsertBreakLate(object doc, int breakType)
        {
            object content = GetProp(doc, "Content");
            int end = Convert.ToInt32(GetProp(content, "End"));
            int pos = end > 1 ? end - 1 : end;
            object range = Invoke(doc, "Range", pos, pos);
            Invoke(range, "InsertBreak", breakType);
        }

        private static bool FindAllStoredTextsLate(object doc, IList<string> storedTexts, out string miss)
        {
            miss = null;
            foreach (string stored in storedTexts)
            {
                if (string.IsNullOrEmpty(stored))
                {
                    continue;
                }

                string converted = WordRangeFinder.ConvertNewlinesToWordCodes(stored);
                object content = Invoke(doc, "Content");
                object find = GetProp(content, "Find");
                Invoke(find, "ClearFormatting");
                SetProp(find, "Text", converted);
                SetProp(find, "MatchCase", true);
                SetProp(find, "Wrap", 0);
                bool hit = false;
                try
                {
                    hit = Convert.ToBoolean(Invoke(find, "Execute"));
                }
                catch
                {
                    hit = false;
                }

                if (!hit)
                {
                    miss = TextUtils.EscapeNewlinesForDb(stored) + " → " + converted;
                    return false;
                }
            }

            return true;
        }

        private static bool InsertPageBreakWord(string fullPath)
        {
            try
            {
                var app = (Word.Application)Marshal.GetActiveObject("Word.Application");
                Word.Document doc = FindWordDoc(app, fullPath);
                if (doc == null)
                {
                    return false;
                }

                object brk = Word.WdBreakType.wdPageBreak;
                doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertBreak(ref brk);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool CloseWordByPath(string fullPath)
        {
            try
            {
                var app = (Word.Application)Marshal.GetActiveObject("Word.Application");
                Word.Document doc = FindWordDoc(app, fullPath);
                if (doc == null)
                {
                    return false;
                }

                object save = false;
                doc.Close(ref save);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static Word.Document FindWordDoc(Word.Application app, string fullPath)
        {
            string norm = Normalize(fullPath);
            foreach (Word.Document d in app.Documents)
            {
                try
                {
                    if (string.Equals(Normalize(d.FullName), norm, StringComparison.OrdinalIgnoreCase))
                    {
                        return d;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private static bool InsertPageBreakWps(string fullPath)
        {
            return WithWpsDoc(fullPath, doc =>
            {
                try
                {
                    object content = Invoke(doc, "Content");
                    int end = Convert.ToInt32(GetProp(content, "End"));
                    int pos = end > 1 ? end - 1 : end;
                    object range = Invoke(doc, "Range", pos, pos);
                    Invoke(range, "InsertBreak", 7);
                    return true;
                }
                catch
                {
                }

                try
                {
                    object content = GetProp(doc, "Content");
                    Invoke(content, "InsertAfter", "\f");
                    return true;
                }
                catch
                {
                }

                return false;
            });
        }

        private static bool CloseWpsByPath(string fullPath)
        {
            return WithWpsDoc(fullPath, doc =>
            {
                Invoke(doc, "Close", false);
                return true;
            });
        }

        private static bool WithWpsDoc(string fullPath, Func<object, bool> action)
        {
            foreach (object app in EnumerateWpsApps())
            {
                object doc = FindLateDoc(app, fullPath);
                if (doc == null)
                {
                    continue;
                }

                try
                {
                    if (action(doc))
                    {
                        return true;
                    }
                }
                catch
                {
                }
            }

            return false;
        }

        private static IEnumerable<object> EnumerateWpsApps()
        {
            var apps = new List<object>();
            string[] ids = { "Kwps.Application", "kwps.Application", "wps.Application" };
            foreach (string id in ids)
            {
                try
                {
                    object app = Marshal.GetActiveObject(id);
                    if (app != null)
                    {
                        apps.Add(app);
                    }
                }
                catch
                {
                }
            }

            CollectWpsAppsFromWindows(apps);
            return apps;
        }

        private static void CollectWpsAppsFromWindows(List<object> apps)
        {
            var pids = new HashSet<uint>();
            try
            {
                foreach (Process p in Process.GetProcessesByName("wps"))
                {
                    try
                    {
                        pids.Add((uint)p.Id);
                    }
                    finally
                    {
                        p.Dispose();
                    }
                }
            }
            catch
            {
                return;
            }

            if (pids.Count == 0)
            {
                return;
            }

            EnumWindows(
                (hWnd, _) =>
                {
                    if (!IsWindowVisible(hWnd))
                    {
                        return true;
                    }

                    GetWindowThreadProcessId(hWnd, out uint pid);
                    if (!pids.Contains(pid))
                    {
                        return true;
                    }

                    TryAddWpsAppFromWindow(hWnd, apps);
                    return true;
                },
                IntPtr.Zero);
        }

        private static void TryAddWpsAppFromWindow(IntPtr hwnd, List<object> apps)
        {
            Guid iid = IidIDispatch;
            object obj;
            if (AccessibleObjectFromWindow(hwnd, ObjIdNativeOm, ref iid, out obj) != 0 || obj == null)
            {
                return;
            }

            try
            {
                if (GetProp(obj, "Documents") != null)
                {
                    apps.Add(obj);
                    return;
                }
            }
            catch
            {
            }

            try
            {
                object app = GetProp(obj, "Application");
                if (app != null)
                {
                    apps.Add(app);
                }
            }
            catch
            {
            }
        }

        private static object FindLateDoc(object app, string fullPath)
        {
            try
            {
                object docs = GetProp(app, "Documents");
                int n = Convert.ToInt32(GetProp(docs, "Count"));
                string wantFile = Path.GetFileName(Normalize(fullPath));
                for (int i = 1; i <= n; i++)
                {
                    object d = Invoke(docs, "Item", i);
                    string name = Convert.ToString(GetProp(d, "FullName") ?? "");
                    if (PathMatches(name, fullPath))
                    {
                        return d;
                    }

                    if (!string.IsNullOrEmpty(wantFile)
                        && string.Equals(
                            Path.GetFileName(Normalize(name)),
                            wantFile,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return d;
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static bool PathMatches(string a, string b)
        {
            string na = Normalize(a);
            string nb = Normalize(b);
            return !string.IsNullOrEmpty(na)
                && !string.IsNullOrEmpty(nb)
                && string.Equals(na, nb, StringComparison.OrdinalIgnoreCase);
        }

        private const uint ObjIdNativeOm = 0xFFFFFFF0;
        private static readonly Guid IidIDispatch = new Guid("00020400-0000-0000-C000-000000000046");

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
            IntPtr hwnd,
            uint dwObjectID,
            ref Guid riid,
            [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

        private static bool WithWordDoc(string fullPath, Func<Word.Document, bool> action)
        {
            try
            {
                var app = (Word.Application)Marshal.GetActiveObject("Word.Application");
                Word.Document doc = FindWordDoc(app, fullPath);
                if (doc == null)
                {
                    return false;
                }

                return action(doc);
            }
            catch
            {
                return false;
            }
        }

        private static void FillWordTable(Word.Table table, string[] cells)
        {
            int n = cells == null ? 0 : cells.Length;
            int k = 0;
            int rows = table.Rows.Count;
            int cols = table.Columns.Count;
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    string t = k < n ? (cells[k] ?? "") : "";
                    k++;
                    table.Cell(r, c).Range.Text = t;
                }
            }
        }

        private static bool InsertTableLate(object doc, int rows, int cols, string[] cells)
        {
            object content = GetProp(doc, "Content");
            int end = Convert.ToInt32(GetProp(content, "End"));
            int pos = end > 0 ? end - 1 : 0;
            object range = Invoke(doc, "Range", pos, pos);
            object tables = GetProp(doc, "Tables");
            object table = Invoke(tables, "Add", range, rows, cols);
            int n = cells == null ? 0 : cells.Length;
            int k = 0;
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    string t = k < n ? (cells[k] ?? "") : "";
                    k++;
                    object cell = Invoke(table, "Cell", r, c);
                    object cr = GetProp(cell, "Range");
                    SetProp(cr, "Text", t);
                }
            }

            return true;
        }

        private static string WriteOnePxPng()
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "ew-wc-" + Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(
                path,
                Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
            return path;
        }

        private static object ParagraphRangeLate(object doc, int paraIndex)
        {
            try
            {
                object paras = GetProp(doc, "Paragraphs");
                int n = Convert.ToInt32(GetProp(paras, "Count"));
                if (paraIndex < 1 || paraIndex > n)
                {
                    return null;
                }

                object para = Invoke(paras, "Item", paraIndex);
                return GetProp(para, "Range");
            }
            catch
            {
                return null;
            }
        }

        private static void ApplyFontLate(object range, string fontName, float size, bool bold)
        {
            object font = GetProp(range, "Font");
            if (!string.IsNullOrEmpty(fontName))
            {
                try
                {
                    SetProp(font, "NameFarEast", fontName);
                }
                catch
                {
                }

                SetProp(font, "Name", fontName);
            }

            if (size > 0)
            {
                SetProp(font, "Size", size);
            }

            SetProp(font, "Bold", bold ? 1 : 0);
        }

        private static object GetProp(object target, string name)
        {
            return target.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.GetProperty,
                null,
                target,
                null);
        }

        private static void SetProp(object target, string name, object value)
        {
            target.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.SetProperty,
                null,
                target,
                new object[] { value });
        }

        private static object Invoke(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                target,
                args);
        }

        private static string Normalize(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return "";
            }

            try
            {
                return Path.GetFullPath(path).TrimEnd('\\').ToLowerInvariant();
            }
            catch
            {
                return path.Trim().ToLowerInvariant();
            }
        }
    }
}
