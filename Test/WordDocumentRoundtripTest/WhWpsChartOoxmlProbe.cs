using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 探 WPS 文字：另存包是否有 chart 部件、InsertFile 是否活图、种子稿不 AddChart 能否读出 &lt;Chart&gt;。
    /// 只动临时文件，不 Quit WPS / Word。
    /// </summary>
    internal static class WhWpsChartOoxmlProbe
    {
        public static int Run()
        {
            var s = new WdSession("wps");
            WdSession.SweepBeforeRun("wps");
            int fail = 0;
            try
            {
                string seedRaw = s.CreateChartFixtureDocx(1);
                if (string.IsNullOrEmpty(seedRaw) || !File.Exists(seedRaw))
                {
                    Console.WriteLine("FAIL 用 Word 做种子稿失败（不调 WPS AddChart）");
                    return 1;
                }

                string seed = Path.Combine(
                    Path.GetTempPath(),
                    "ew-wd-seed-keep-" + Guid.NewGuid().ToString("N") + ".docx");
                File.Copy(seedRaw, seed, true);
                Line("种子", seed);
                Line("种子(Word落盘)", seedRaw);
                fail += Section("1 Word 种子包", () => DumpChartParts(seed, "Word种子"));

                string saveAsPath = s.NewDocxPath();
                fail += Section("2 WPS 另存包", () => ProbeSaveAsPackage(s, seed, saveAsPath));
                fail += Section("2b 直接打开种子", () => ProbeOpenExisting(s, seed, "Word种子"));
                fail += Section("2c 直接打开WPS另存", () =>
                    File.Exists(saveAsPath)
                        ? ProbeOpenExisting(s, saveAsPath, "WPS另存")
                        : FailLine("WPS另存文件不在"));

                string insertPath = s.NewDocxPath();
                string insertId = null;
                fail += Section("3 COM InsertFile", () => ProbeInsertFile(s, seed, insertPath, out insertId, comInsert: true));
                fail += Section("3b 四件工具", () => ProbeFourTools(s, insertId));
                fail += Section("3c COM 改 Series", () => ProbeWriteSeries(s, insertPath));

                string toolInsertPath = s.NewDocxPath();
                string toolInsertId = null;
                fail += Section("3d 产品 InsertFile", () => ProbeInsertFile(s, seed, toolInsertPath, out toolInsertId, comInsert: false));

                if (!LooksLive(s, insertPath) && !LooksLive(s, toolInsertPath))
                {
                    fail += Section("4 跨文档 Range Copy", () => ProbeRangeCopy(s, seed));
                }
                else
                {
                    Console.WriteLine("  跳过 4：InsertFile 已是活图");
                }

                // 改包后再 Open 容易把 WPS COM 弄死，本趟先不跑。另存包已证明有 chart1.xml。

                return fail == 0 ? 0 : 1;
            }
            finally
            {
                WdSession.SweepAfterRun("wps");
            }
        }

        private static int Section(string title, Func<int> body)
        {
            Console.WriteLine();
            Console.WriteLine("==== " + title + " ====");
            try
            {
                return body();
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL 异常: " + ex.Message);
                return 1;
            }
        }

        private static int ProbeSaveAsPackage(WdSession s, string seed, string dest)
        {
            string openCopy = Path.Combine(
                Path.GetTempPath(),
                "ew-wd-seed-open-" + Guid.NewGuid().ToString("N") + ".docx");
            File.Copy(seed, openCopy, true);
            ToolResult opened = s.Open(openCopy, false, "wps");
            if (!opened.Success)
            {
                Line("打开种子副本", opened.Error ?? "失败");
                return 1;
            }

            bool saved = false;
            string saveHow = "";
            s.TryWpsDoc(openCopy, doc =>
            {
                try
                {
                    Invoke(doc, "SaveCopyAs", dest);
                    saved = File.Exists(dest);
                    saveHow = "SaveCopyAs";
                    return saved;
                }
                catch (Exception ex)
                {
                    saveHow = "SaveCopyAs 无/失败: " + TrimErr(ex.Message);
                }

                try
                {
                    Invoke(doc, "SaveAs", dest);
                    saved = File.Exists(dest);
                    saveHow = "SaveAs";
                    return saved;
                }
                catch (Exception ex)
                {
                    saveHow = saveHow + " | SaveAs 失败: " + TrimErr(ex.Message);
                    return false;
                }
            });

            Line("另存方式", saveHow);
            Line("另存路径", dest);
            Line("另存成功", saved ? "是" : "否");
            if (!saved)
            {
                s.CloseDocByPath(seed);
                return 1;
            }

            string readPath = dest;
            string copy = dest + ".copy.docx";
            try
            {
                File.Copy(dest, copy, true);
                readPath = copy;
            }
            catch (Exception ex)
            {
                Line("另存副本(解锁读)", TrimErr(ex.Message));
            }

            int n = DumpChartParts(readPath, "WPS另存");
            string afterName = "";
            s.TryWpsDoc(openCopy, doc =>
            {
                afterName = Convert.ToString(GetProp(doc, "FullName") ?? "");
                return true;
            });
            if (string.IsNullOrEmpty(afterName))
            {
                s.TryWpsDoc(dest, doc =>
                {
                    afterName = Convert.ToString(GetProp(doc, "FullName") ?? "");
                    return true;
                });
            }

            Line("另存后 FullName", afterName);
            Line("原种子还在", File.Exists(seed) ? "是" : "否");
            Line("打开副本还在", File.Exists(openCopy) ? "是" : "否");
            bool wroteUserSeed = PathsEqual(afterName, seed) || PathsEqual(afterName, openCopy);
            Line("是否写回打开路径", wroteUserSeed ? "是（SaveAs 会改 FullName）" : "否");
            s.CloseDocByPath(openCopy);
            s.CloseDocByPath(dest);
            return n;
        }

        private static int FailLine(string message)
        {
            Line("结果", message);
            return 1;
        }

        private static int ProbeOpenExisting(WdSession s, string path, string tag)
        {
            string openCopy = Path.Combine(
                Path.GetTempPath(),
                "ew-wd-open-" + Guid.NewGuid().ToString("N") + ".docx");
            File.Copy(path, openCopy, true);
            ToolResult opened = s.Open(openCopy, false, "wps");
            if (!opened.Success)
            {
                Line(tag + " 打开", opened.Error ?? "失败");
                return 1;
            }

            string id = WdSession.Str(opened, "channel_id");
            string dump = DescribeCharts(s, openCopy);
            Line(tag + " 形状", dump);
            Line(tag + " 活图", LooksLiveDump(dump) ? "是" : "否");
            ToolResult got = s.GetContent(id);
            string body = ReadBody(got);
            bool hasChartTag = (body ?? "").IndexOf("<Chart", StringComparison.OrdinalIgnoreCase) >= 0;
            Line(tag + " 读文 <Chart", hasChartTag ? "是" : "否");
            Line(tag + " 读文片段", Snip(body, 280));
            Line(tag + " C_", FindChartId(body) ?? "无");
            ProbeWriteSeries(s, openCopy);
            return (LooksLiveDump(dump) || hasChartTag) ? 0 : 1;
        }

        private static int ProbeInsertFile(
            WdSession s,
            string seed,
            string blankPath,
            out string channelId,
            bool comInsert)
        {
            channelId = null;
            ToolResult opened = s.Open(blankPath, true, "wps");
            if (!opened.Success)
            {
                Line("空稿", opened.Error ?? "失败");
                return 1;
            }

            channelId = WdSession.Str(opened, "channel_id");
            Line("空稿 channel", channelId ?? "");
            bool inserted;
            if (comInsert)
            {
                inserted = s.TryWpsDoc(blankPath, doc => TryInsertSeed(doc, seed));
                Line("COM 插入种子", inserted ? "成功" : "失败");
            }
            else
            {
                ToolResult ins = s.InsertFromFile(new Dictionary<string, object>
                {
                    ["channel_id"] = channelId,
                    ["path"] = seed,
                    ["format"] = new Dictionary<string, object>
                    {
                        ["char_format"] = new Dictionary<string, object>
                        {
                            ["font_name"] = "宋体",
                            ["font_size"] = 12,
                            ["is_bold"] = false
                        }
                    }
                });
                inserted = ins.Success;
                Line("F_insert_document_from_file", inserted ? "成功" : (ins.Error ?? "失败"));
            }

            if (!inserted)
            {
                return 1;
            }

            string dump = DescribeCharts(s, blankPath);
            Line("形状", dump);
            bool live = LooksLiveDump(dump);
            Line("活图", live ? "是" : "否");

            ToolResult got = s.GetContent(channelId);
            string body = ReadBody(got);
            bool hasChartTag = (body ?? "").IndexOf("<Chart", StringComparison.OrdinalIgnoreCase) >= 0;
            Line("读文 Success", got.Success ? "是" : (got.Error ?? "否"));
            Line("读文含 <Chart", hasChartTag ? "是" : "否");
            Line("读文片段", Snip(body, 280));
            return (live || hasChartTag) ? 0 : 1;
        }

        private static int ProbeFourTools(WdSession s, string channelId)
        {
            if (string.IsNullOrEmpty(channelId))
            {
                Line("四件", "无 channel，跳过");
                return 0;
            }

            ToolResult got = s.GetContent(channelId);
            string body = ReadBody(got);
            string chartId = FindChartId(body);
            Line("读文 C_", chartId ?? "无");

            var dummyXml = Path.Combine(Path.GetTempPath(), "ew-wd-probe-" + Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(dummyXml, "<ChartConfig><Chart><Type>column</Type></Chart></ChartConfig>", Encoding.UTF8);

            ToolResult create = s.CreateChartFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["path"] = dummyXml
            });
            Line("create", DescribeTool(create));

            ToolResult extract = s.ExtractChartXml(new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["chart_selection"] = new Dictionary<string, object>
                {
                    ["type"] = "by_id",
                    ["chart_id"] = chartId ?? "C_00000000"
                },
                ["scope"] = "all"
            });
            Line("extract", DescribeTool(extract));

            ToolResult apply = s.ApplyChartFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["chart_id"] = chartId ?? "C_00000000",
                ["scope"] = "format",
                ["xml_content"] = "<ChartConfig><Chart><Type>column</Type><Title>易写探针套</Title></Chart></ChartConfig>"
            });
            Line("apply", DescribeTool(apply));

            ToolResult del = s.DeleteChart(new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["chart_id"] = chartId ?? "C_00000000"
            });
            Line("delete", DescribeTool(del));
            return 0;
        }

        private static int ProbeWriteSeries(WdSession s, string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath))
            {
                return 0;
            }

            bool wrote = false;
            string detail = "";
            s.TryWpsDoc(fullPath, doc =>
            {
                object chart = FirstChart(doc, out detail);
                if (chart == null)
                {
                    return false;
                }

                try
                {
                    object sc = Invoke(chart, "SeriesCollection");
                    object one = Invoke(sc, "Item", 1);
                    if (one == null)
                    {
                        one = Invoke(chart, "SeriesCollection", 1);
                    }

                    SetProp(one, "Values", new object[] { 11, 22, 33 });
                    object vals = GetProp(one, "Values");
                    detail = "Values 写后=" + FormatVal(vals);
                    wrote = true;
                    return true;
                }
                catch (Exception ex)
                {
                    detail = "写 Values 失败: " + TrimErr(ex.Message);
                    return false;
                }
            });
            Line("COM 套数", wrote ? detail : (detail == "" ? "无图" : detail));
            return 0;
        }

        private static int ProbeRangeCopy(WdSession s, string seed)
        {
            ToolResult src = s.Open(seed, false, "wps");
            if (!src.Success)
            {
                Line("开种子", src.Error ?? "失败");
                return 1;
            }

            string destPath = s.NewDocxPath();
            ToolResult dest = s.Open(destPath, true, "wps");
            if (!dest.Success)
            {
                Line("开空稿", dest.Error ?? "失败");
                s.CloseDocByPath(seed);
                return 1;
            }

            bool copied = false;
            string err = "";
            s.TryWpsDoc(seed, srcDoc =>
            {
                object shape = FirstInline(srcDoc);
                if (shape == null)
                {
                    err = "种子无 InlineShape";
                    return false;
                }

                try
                {
                    object range = GetProp(shape, "Range");
                    Invoke(range, "Copy");
                }
                catch (Exception ex)
                {
                    err = "Copy 失败: " + TrimErr(ex.Message);
                    return false;
                }

                return s.TryWpsDoc(destPath, destDoc =>
                {
                    try
                    {
                        object content = GetProp(destDoc, "Content");
                        Invoke(content, "Paste");
                        copied = true;
                        return true;
                    }
                    catch (Exception ex)
                    {
                        err = "Paste 失败: " + TrimErr(ex.Message);
                        return false;
                    }
                });
            });

            Line("Range Copy/Paste", copied ? "成功" : err);
            Line("粘贴后形状", DescribeCharts(s, destPath));
            string destId = WdSession.Str(dest, "channel_id");
            ToolResult got = s.GetContent(destId);
            string body = ReadBody(got);
            Line("粘贴后读文 <Chart", (body ?? "").IndexOf("<Chart", StringComparison.OrdinalIgnoreCase) >= 0 ? "是" : "否");
            Line("粘贴后读文片段", Snip(body, 220));
            s.CloseDocByPath(seed);
            return copied ? 0 : 1;
        }

        private static int DumpChartParts(string path, string tag)
        {
            if (!File.Exists(path))
            {
                Line(tag, "文件不存在");
                return 1;
            }

            long size = new FileInfo(path).Length;
            Line(tag + " 大小", size + " B");
            byte[] head = new byte[4];
            using (var fs = File.OpenRead(path))
            {
                fs.Read(head, 0, 4);
            }

            bool zip = head[0] == 0x50 && head[1] == 0x4B;
            Line(tag + " 是 zip", zip ? "是" : ("否 head=" + BitConverter.ToString(head)));
            if (!zip)
            {
                return 1;
            }

            var charts = new List<string>();
            var embeds = new List<string>();
            var other = new List<string>();
            try
            {
                using (ZipArchive z = ZipFile.OpenRead(path))
                {
                    foreach (ZipArchiveEntry e in z.Entries)
                    {
                        string n = e.FullName.Replace('\\', '/');
                        if (n.IndexOf("word/charts/", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            charts.Add(n + " (" + e.Length + ")");
                        }
                        else if (n.IndexOf("word/embeddings/", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            embeds.Add(n + " (" + e.Length + ")");
                        }
                        else if (n.IndexOf("chart", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            other.Add(n + " (" + e.Length + ")");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Line(tag + " 解包", TrimErr(ex.Message));
                return 1;
            }

            Line(tag + " word/charts", charts.Count == 0 ? "无" : string.Join(" | ", charts));
            Line(tag + " embeddings", embeds.Count == 0 ? "无" : string.Join(" | ", embeds));
            if (other.Count > 0)
            {
                Line(tag + " 其它含 chart", string.Join(" | ", other));
            }

            bool hasChart1 = charts.Exists(x =>
                x.IndexOf("word/charts/chart1.xml", StringComparison.OrdinalIgnoreCase) >= 0
                || x.IndexOf("word/charts/chart", StringComparison.OrdinalIgnoreCase) >= 0);
            return hasChart1 ? 0 : 1;
        }

        private static string RewriteSeedSeriesName(string seed, string from, string to)
        {
            string dest = Path.Combine(
                Path.GetTempPath(),
                "ew-wd-seed-rw-" + Guid.NewGuid().ToString("N") + ".docx");
            File.Copy(seed, dest, true);
            try
            {
                using (ZipArchive z = ZipFile.Open(dest, ZipArchiveMode.Update))
                {
                    var names = new List<string>();
                    foreach (ZipArchiveEntry e in z.Entries)
                    {
                        names.Add(e.FullName);
                    }

                    foreach (string name in names)
                    {
                        string n = name.Replace('\\', '/');
                        bool xml = n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
                        bool xlsx = n.IndexOf("word/embeddings/", StringComparison.OrdinalIgnoreCase) >= 0
                            && n.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);
                        if (!xml && !xlsx)
                        {
                            continue;
                        }

                        ZipArchiveEntry entry = z.GetEntry(name);
                        if (entry == null)
                        {
                            continue;
                        }

                        byte[] raw;
                        using (Stream input = entry.Open())
                        using (var ms = new MemoryStream())
                        {
                            input.CopyTo(ms);
                            raw = ms.ToArray();
                        }

                        if (xml)
                        {
                            string text = Encoding.UTF8.GetString(raw);
                            if (text.IndexOf(from, StringComparison.Ordinal) < 0)
                            {
                                continue;
                            }

                            text = text.Replace(from, to);
                            entry.Delete();
                            ZipArchiveEntry next = z.CreateEntry(name, CompressionLevel.Optimal);
                            using (Stream output = next.Open())
                            {
                                byte[] bytes = Encoding.UTF8.GetBytes(text);
                                output.Write(bytes, 0, bytes.Length);
                            }
                        }
                        else
                        {
                            byte[] nextXlsx = ReplaceInXlsx(raw, from, to);
                            if (nextXlsx == null)
                            {
                                continue;
                            }

                            entry.Delete();
                            ZipArchiveEntry next = z.CreateEntry(name, CompressionLevel.Optimal);
                            using (Stream output = next.Open())
                            {
                                output.Write(nextXlsx, 0, nextXlsx.Length);
                            }
                        }
                    }
                }

                return dest;
            }
            catch (Exception ex)
            {
                Console.WriteLine("  改包失败: " + TrimErr(ex.Message));
                return null;
            }
        }

        private static byte[] ReplaceInXlsx(byte[] xlsx, string from, string to)
        {
            using (var input = new MemoryStream(xlsx))
            using (var z = new ZipArchive(input, ZipArchiveMode.Read))
            using (var output = new MemoryStream())
            {
                using (var zo = new ZipArchive(output, ZipArchiveMode.Create, true))
                {
                    bool any = false;
                    foreach (ZipArchiveEntry e in z.Entries)
                    {
                        byte[] raw;
                        using (Stream s = e.Open())
                        using (var ms = new MemoryStream())
                        {
                            s.CopyTo(ms);
                            raw = ms.ToArray();
                        }

                        if (e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                        {
                            string text = Encoding.UTF8.GetString(raw);
                            if (text.IndexOf(from, StringComparison.Ordinal) >= 0)
                            {
                                text = text.Replace(from, to);
                                raw = Encoding.UTF8.GetBytes(text);
                                any = true;
                            }
                        }

                        ZipArchiveEntry n = zo.CreateEntry(e.FullName, CompressionLevel.Optimal);
                        using (Stream o = n.Open())
                        {
                            o.Write(raw, 0, raw.Length);
                        }
                    }

                    if (!any)
                    {
                        return null;
                    }
                }

                return output.ToArray();
            }
        }

        private static bool TryInsertSeed(object destDoc, string seed)
        {
            TrySilenceAlerts(destDoc);
            object content = GetProp(destDoc, "Content");
            object app = GetProp(destDoc, "Application");
            object sel = app == null ? null : GetProp(app, "Selection");
            object missing = Type.Missing;

            if (TryCall("InsertFile(content,path)", () =>
            {
                Invoke(content, "InsertFile", seed);
            }))
            {
                return true;
            }

            if (TryCall("InsertFile(content,path,missing,false)", () =>
            {
                Invoke(content, "InsertFile", seed, missing, false);
            }))
            {
                return true;
            }

            if (sel != null && TryCall("InsertFile(sel,path,missing,false,false)", () =>
            {
                Invoke(sel, "InsertFile", seed, missing, false, false);
            }))
            {
                return true;
            }

            object documents = app == null ? null : GetProp(app, "Documents");
            if (documents != null)
            {
                object src = null;
                try
                {
                    src = Invoke(documents, "Open", seed);
                    object srcContent = GetProp(src, "Content");
                    if (TryCall("FormattedText", () =>
                    {
                        object formatted = GetProp(srcContent, "FormattedText");
                        SetProp(content, "FormattedText", formatted);
                    }))
                    {
                        TryCloseDoc(src);
                        return true;
                    }

                    if (TryCall("Range.Copy + Content.Paste", () =>
                    {
                        Invoke(srcContent, "Copy");
                        Invoke(content, "Paste");
                    }))
                    {
                        TryCloseDoc(src);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Line("开种子作源", TrimErr(ex.Message));
                }

                TryCloseDoc(src);
            }

            return false;
        }

        private static bool TryCall(string tag, Action action)
        {
            try
            {
                action();
                Line(tag, "成功");
                return true;
            }
            catch (Exception ex)
            {
                Line(tag, TrimErr(ex.Message));
                return false;
            }
        }

        private static void TrySilenceAlerts(object doc)
        {
            try
            {
                object app = GetProp(doc, "Application");
                SetProp(app, "DisplayAlerts", 0);
            }
            catch
            {
            }
        }

        private static void TryCloseDoc(object doc)
        {
            if (doc == null)
            {
                return;
            }

            try
            {
                Invoke(doc, "Close", false);
            }
            catch
            {
            }
        }

        private static bool LooksLive(WdSession s, string path)
        {
            return LooksLiveDump(DescribeCharts(s, path));
        }

        private static bool LooksLiveDump(string dump)
        {
            if (string.IsNullOrEmpty(dump))
            {
                return false;
            }

            return dump.IndexOf("HasChart=True", StringComparison.OrdinalIgnoreCase) >= 0
                || dump.IndexOf("series=", StringComparison.OrdinalIgnoreCase) >= 0
                || dump.IndexOf("type=12", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string DescribeCharts(WdSession s, string fullPath)
        {
            string dump = "未打开";
            s.TryWpsDoc(fullPath, doc =>
            {
                var sb = new StringBuilder();
                DumpCollection(sb, "InlineShapes", GetProp(doc, "InlineShapes"), true);
                DumpCollection(sb, "Shapes", GetProp(doc, "Shapes"), false);
                dump = sb.Length == 0 ? "无集合" : sb.ToString();
                return true;
            });
            return dump;
        }

        private static void DumpCollection(StringBuilder sb, string label, object collection, bool preferChart)
        {
            if (collection == null)
            {
                sb.Append(label).Append("=null ");
                return;
            }

            int count;
            try
            {
                count = Convert.ToInt32(GetProp(collection, "Count"));
            }
            catch (Exception ex)
            {
                sb.Append(label).Append(" Count失败=").Append(TrimErr(ex.Message)).Append(' ');
                return;
            }

            sb.Append(label).Append("[").Append(count).Append("] ");
            for (int i = 1; i <= count; i++)
            {
                object shape;
                try
                {
                    shape = Invoke(collection, "Item", i);
                }
                catch
                {
                    continue;
                }

                int type = -1;
                try
                {
                    type = Convert.ToInt32(GetProp(shape, "Type"));
                }
                catch
                {
                }

                object has = null;
                try
                {
                    has = GetProp(shape, "HasChart");
                }
                catch
                {
                }

                string series = "";
                object chart = TryChart(shape);
                if (chart != null)
                {
                    try
                    {
                        object sc = Invoke(chart, "SeriesCollection");
                        int n = Convert.ToInt32(GetProp(sc, "Count"));
                        object one = Invoke(sc, "Item", 1);
                        string name = Convert.ToString(GetProp(one, "Name"));
                        object vals = GetProp(one, "Values");
                        series = " series=" + n + " name=" + name + " values=" + FormatVal(vals);
                    }
                    catch (Exception ex)
                    {
                        series = " Chart在 Series失败=" + TrimErr(ex.Message);
                    }
                }

                sb.Append(label).Append("#").Append(i)
                    .Append(" type=").Append(type)
                    .Append(" HasChart=").Append(has ?? "?")
                    .Append(chart == null ? " 无Chart" : " 有Chart")
                    .Append(series)
                    .Append("; ");
            }
        }

        private static object FirstChart(object doc, out string detail)
        {
            detail = "";
            object shape = FirstInline(doc);
            if (shape == null)
            {
                detail = "无 InlineShape";
                return null;
            }

            object chart = TryChart(shape);
            if (chart == null)
            {
                detail = "InlineShape 无 Chart";
            }

            return chart;
        }

        private static object FirstInline(object doc)
        {
            object inlines = GetProp(doc, "InlineShapes");
            int n = Convert.ToInt32(GetProp(inlines, "Count"));
            if (n < 1)
            {
                return null;
            }

            return Invoke(inlines, "Item", 1);
        }

        private static object TryChart(object shape)
        {
            try
            {
                object has = GetProp(shape, "HasChart");
                if (has is bool b && !b)
                {
                    return null;
                }
            }
            catch
            {
            }

            try
            {
                object chart = GetProp(shape, "Chart");
                if (chart != null)
                {
                    return chart;
                }
            }
            catch
            {
            }

            return null;
        }

        private static string ReadBody(ToolResult got)
        {
            if (got == null)
            {
                return "";
            }

            foreach (string key in new[] { "document_content", "text", "content" })
            {
                string v = WdSession.Str(got, key);
                if (!string.IsNullOrEmpty(v))
                {
                    return v;
                }
            }

            return Blob(got);
        }

        private static string FindChartId(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return null;
            }

            int i = body.IndexOf("id=\"C_", StringComparison.Ordinal);
            if (i < 0)
            {
                i = body.IndexOf("id='C_", StringComparison.Ordinal);
            }

            if (i < 0)
            {
                return null;
            }

            int start = body.IndexOf("C_", i, StringComparison.Ordinal);
            if (start < 0)
            {
                return null;
            }

            int end = start + 2;
            while (end < body.Length && (char.IsLetterOrDigit(body[end])))
            {
                end++;
            }

            return body.Substring(start, end - start);
        }

        private static string DescribeTool(ToolResult r)
        {
            if (r == null)
            {
                return "null";
            }

            if (r.Success)
            {
                return "成功 " + Snip(Blob(r), 120);
            }

            string err = r.Error ?? "";
            bool uns = err.IndexOf("unsupported", StringComparison.OrdinalIgnoreCase) >= 0;
            return (uns ? "unsupported " : "失败 ") + Snip(err, 160);
        }

        private static string FormatVal(object vals)
        {
            if (vals == null)
            {
                return "null";
            }

            if (vals is Array arr)
            {
                var parts = new List<string>();
                foreach (object x in arr)
                {
                    parts.Add(Convert.ToString(x));
                    if (parts.Count >= 6)
                    {
                        break;
                    }
                }

                return "[" + string.Join(",", parts) + "]";
            }

            return Convert.ToString(vals);
        }

        private static string Blob(ToolResult r)
        {
            if (r == null)
            {
                return "";
            }

            try
            {
                return Newtonsoft.Json.JsonConvert.SerializeObject(r);
            }
            catch
            {
                return r.Error ?? "";
            }
        }

        private static void Line(string k, string v)
        {
            Console.WriteLine("  " + k + ": " + (v ?? ""));
        }

        private static string Snip(string text, int max)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string one = text.Replace("\r", " ").Replace("\n", " ");
            return one.Length <= max ? one : one.Substring(0, max) + "…";
        }

        private static string TrimErr(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return "";
            }

            return message.Replace("\r", " ").Replace("\n", " ");
        }

        private static bool PathsEqual(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
            {
                return false;
            }

            try
            {
                return string.Equals(
                    Path.GetFullPath(a).TrimEnd('\\'),
                    Path.GetFullPath(b).TrimEnd('\\'),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }

        private static object GetProp(object target, string name)
        {
            return target.GetType().InvokeMember(
                name,
                BindingFlags.GetProperty,
                null,
                target,
                null);
        }

        private static void SetProp(object target, string name, object value)
        {
            target.GetType().InvokeMember(
                name,
                BindingFlags.SetProperty,
                null,
                target,
                new object[] { value });
        }

        private static object Invoke(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(
                name,
                BindingFlags.InvokeMethod,
                null,
                target,
                args);
        }
    }
}
