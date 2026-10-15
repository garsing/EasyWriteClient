using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 探 WPS 文字：另存包是否有 chart 部件、InsertFile 是否活图、种子稿不 AddChart 能否读出 &lt;Chart&gt;。
    /// 只动临时文件，不 Quit WPS / Word。
    /// </summary>
    internal static class WhWpsChartOoxmlProbe
    {
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr extra);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT point);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWnd, EnumChildProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

        [DllImport("user32.dll")]
        private static extern bool ScreenToClient(IntPtr hWnd, ref POINT point);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        private delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        private const int SwRestore = 9;
        private const uint GaRoot = 2;
        private const uint MouseeventfLeftdown = 0x0002;
        private const uint MouseeventfLeftup = 0x0004;
        private const uint WmMousemove = 0x0200;
        private const uint WmLbuttondown = 0x0201;
        private const uint WmLbuttonup = 0x0202;
        private const uint KeyeventfKeyup = 0x0002;
        private const byte VkMenu = 0x12;

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public static int RunOoxmlReadWhenSeriesDead()
        {
            var s = new WdSession("wps");
            string src = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-src.docx");
            string blank = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-blank.docx");
            if (!File.Exists(src) || !File.Exists(blank))
            {
                Console.WriteLine("找不到样例稿");
                return 1;
            }

            ToolResult openSrc = s.Open(src, false, "wps");
            ToolResult openBlank = s.Open(blank, false, "wps");
            if (!openSrc.Success || !openBlank.Success)
            {
                Console.WriteLine("打开失败 src=" + (openSrc.Error ?? "") + " blank=" + (openBlank.Error ?? ""));
                return 1;
            }

            int deleted = 0;
            s.TryWpsDoc(blank, destDoc =>
            {
                deleted = DeleteInlineShapes(destDoc);
                return true;
            });
            Line("先清空稿图", deleted.ToString());

            bool copied = PasteSrcChartToBlank(s, src, blank, out string err);
            if (!copied)
            {
                Console.WriteLine("FAIL 复制失败: " + err);
                return 1;
            }

            string comDump = DescribeCharts(s, blank);
            Line("COM Series", comDump);
            bool comLive = LooksLiveDump(comDump);
            Line("COM 活图", comLive ? "是（本趟没法证明废图旁路）" : "否");

            string channelId = WdSession.Str(openBlank, "channel_id");
            ToolResult got = s.GetContent(channelId);
            string body = ReadBody(got);
            Line("读文 <Chart", (body ?? "").IndexOf("<Chart", StringComparison.OrdinalIgnoreCase) >= 0 ? "是" : "否");
            Line("读文片段", Snip(body, 280));

            string snap = SnapshotWithoutStealPath(s, blank);
            if (string.IsNullOrEmpty(snap) || !File.Exists(snap))
            {
                Console.WriteLine("FAIL 副本没落盘");
                s.ActivateDoc(blank);
                return 1;
            }

            Line("副本", snap);
            DumpChartParts(snap, "副本包");
            string ooxml = ReadChartSeriesFromDocx(snap);
            Line("OOXML 套数", ooxml);
            bool ooxmlLive = ooxml.IndexOf("values=[", StringComparison.Ordinal) >= 0
                && ooxml.IndexOf("values=[]", StringComparison.Ordinal) < 0;

            string fullName = "";
            s.TryWpsDoc(blank, doc =>
            {
                fullName = Convert.ToString(GetProp(doc, "FullName") ?? "");
                return true;
            });
            Line("窗口 FullName", fullName);
            Line("窗口还在原路径", PathsEqual(fullName, blank) ? "是" : "否");

            Console.WriteLine();
            if (!comLive && ooxmlLive)
            {
                Console.WriteLine("COM Series 废，副本 OOXML 有数。读数可以走旁路。");
            }
            else if (comLive && ooxmlLive)
            {
                Console.WriteLine("副本 OOXML 有数，但本次 COM 也是活的。");
            }
            else
            {
                Console.WriteLine("副本 OOXML 没读到套数，这条旁路还不通。");
            }

            s.ActivateDoc(blank);
            return ooxmlLive ? 0 : 1;
        }

        public static int RunClickWakeAfterPaste()
        {
            try
            {
                SetProcessDPIAware();
            }
            catch
            {
            }

            var s = new WdSession("wps");
            string src = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-src.docx");
            string blank = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-blank.docx");
            if (!File.Exists(src) || !File.Exists(blank))
            {
                Console.WriteLine("找不到样例稿");
                return 1;
            }

            ToolResult openSrc = s.Open(src, false, "wps");
            ToolResult openBlank = s.Open(blank, false, "wps");
            if (!openSrc.Success || !openBlank.Success)
            {
                Console.WriteLine("打开失败 src=" + (openSrc.Error ?? "") + " blank=" + (openBlank.Error ?? ""));
                return 1;
            }

            int deleted = 0;
            s.TryWpsDoc(blank, destDoc =>
            {
                deleted = DeleteInlineShapes(destDoc);
                return true;
            });
            Line("先清空稿图", deleted.ToString());

            bool copied = false;
            string err = "";
            s.TryWpsDoc(src, srcDoc =>
            {
                object shape = FirstInline(srcDoc);
                if (shape == null)
                {
                    err = "带图稿没有 InlineShape";
                    return false;
                }

                try
                {
                    TrySilenceAlerts(srcDoc);
                    object range = GetProp(shape, "Range");
                    Invoke(range, "Select");
                    Invoke(range, "Copy");
                }
                catch (Exception ex)
                {
                    err = "Copy 失败: " + TrimErr(ex.Message);
                    return false;
                }

                return s.TryWpsDoc(blank, destDoc =>
                {
                    TrySilenceAlerts(destDoc);
                    object content = GetProp(destDoc, "Content");
                    try
                    {
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
            if (!copied)
            {
                Console.WriteLine("FAIL 复制失败: " + err);
                return 1;
            }

            string before = DescribeCharts(s, blank);
            Line("贴后立刻", before);
            if (LooksLiveDump(before))
            {
                Console.WriteLine("贴完就已经能读 Series，没法验证点击唤醒。");
                s.ActivateDoc(blank);
                return 0;
            }

            bool live = TryClickWakePoints(s, blank);
            Console.WriteLine();
            Console.WriteLine(live
                ? "模拟点击后 Series 能读，这条唤醒路可走。"
                : "多点了可见区，Series 仍废。窗口不关。");
            s.ActivateDoc(blank);
            return live ? 0 : 1;
        }

        public static int RunWakeAfterPaste()
        {
            var s = new WdSession("wps");
            string src = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-src.docx");
            if (!File.Exists(src))
            {
                string seed = s.CreateChartFixtureDocx(1);
                if (string.IsNullOrEmpty(seed))
                {
                    Console.WriteLine("FAIL 无种子");
                    return 1;
                }

                File.Copy(seed, src, true);
            }

            string dest = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-wake.docx");
            ToolResult openSrc = s.Open(src, false, "wps");
            ToolResult openDest = File.Exists(dest)
                ? s.Open(dest, false, "wps")
                : s.Open(dest, true, "wps");
            if (!openSrc.Success || !openDest.Success)
            {
                Console.WriteLine("打开失败 " + (openSrc.Error ?? "") + " " + (openDest.Error ?? ""));
                return 1;
            }

            bool copied = false;
            s.TryWpsDoc(src, srcDoc =>
            {
                object shape = FirstInline(srcDoc);
                if (shape == null)
                {
                    return false;
                }

                TrySilenceAlerts(srcDoc);
                Invoke(GetProp(shape, "Range"), "Copy");
                return s.TryWpsDoc(dest, destDoc =>
                {
                    TrySilenceAlerts(destDoc);
                    Invoke(GetProp(destDoc, "Content"), "Paste");
                    copied = true;
                    return true;
                });
            });
            if (!copied)
            {
                Console.WriteLine("FAIL 复制失败");
                return 1;
            }

            Line("贴后立刻", DescribeCharts(s, dest));
            TryWake(s, dest, "Select形状");
            Line("Select后", DescribeCharts(s, dest));
            System.Threading.Thread.Sleep(800);
            Line("等800ms", DescribeCharts(s, dest));
            TryWake(s, dest, "Activate Chart");
            Line("Activate后", DescribeCharts(s, dest));
            System.Windows.Forms.Application.DoEvents();
            System.Threading.Thread.Sleep(800);
            Line("DoEvents+再等", DescribeCharts(s, dest));
            TryWake(s, dest, "Chart.Refresh");
            Line("Refresh后", DescribeCharts(s, dest));
            TryWakeChartData(s, dest);
            Line("ChartData后", DescribeCharts(s, dest));
            TryWakeRemarshal(s, dest);
            Line("重取COM后", DescribeCharts(s, dest));

            string last = DescribeCharts(s, dest);
            bool live = LooksLiveDump(last);
            Console.WriteLine();
            Console.WriteLine(live ? "唤醒后 Series 能读，贴回这条路可走。" : "仍读不到 Series，还要再找激活手段。");
            Console.WriteLine("最终: " + last);
            s.ActivateDoc(dest);
            return live ? 0 : 1;
        }

        private static void TryWakeChartData(WdSession s, string path)
        {
            s.TryWpsDoc(path, doc =>
            {
                object chart = TryChart(FirstInline(doc));
                if (chart == null)
                {
                    Line("ChartData", "无 Chart");
                    return true;
                }

                try
                {
                    object cd = GetProp(chart, "ChartData");
                    try
                    {
                        Invoke(cd, "Activate");
                        Line("ChartData.Activate", "成功");
                    }
                    catch (Exception ex)
                    {
                        Line("ChartData.Activate", TrimErr(ex.Message));
                    }

                    try
                    {
                        object wb = GetProp(cd, "Workbook");
                        Line("ChartData.Workbook", wb == null ? "null" : wb.GetType().FullName);
                    }
                    catch (Exception ex)
                    {
                        Line("ChartData.Workbook", TrimErr(ex.Message));
                    }
                }
                catch (Exception ex)
                {
                    Line("ChartData", TrimErr(ex.Message));
                }

                return true;
            });
        }

        private static void TryWakeRemarshal(WdSession s, string path)
        {
            System.Runtime.InteropServices.Marshal.CleanupUnusedObjectsInCurrentContext();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Line("重取COM", s.TryWpsDoc(path, _ => true) ? "文档还在" : "找不到文档");
        }

        private static bool PasteSrcChartToBlank(WdSession s, string src, string blank, out string err)
        {
            err = "";
            bool copied = false;
            string copyErr = "";
            s.TryWpsDoc(src, srcDoc =>
            {
                object shape = FirstInline(srcDoc);
                if (shape == null)
                {
                    copyErr = "带图稿没有 InlineShape";
                    return false;
                }

                try
                {
                    TrySilenceAlerts(srcDoc);
                    object range = GetProp(shape, "Range");
                    Invoke(range, "Select");
                    Invoke(range, "Copy");
                }
                catch (Exception ex)
                {
                    copyErr = "Copy 失败: " + TrimErr(ex.Message);
                    return false;
                }

                return s.TryWpsDoc(blank, destDoc =>
                {
                    TrySilenceAlerts(destDoc);
                    try
                    {
                        Invoke(GetProp(destDoc, "Content"), "Paste");
                        copied = true;
                        return true;
                    }
                    catch (Exception ex)
                    {
                        copyErr = "Paste 失败: " + TrimErr(ex.Message);
                        return false;
                    }
                });
            });
            err = copyErr;
            return copied;
        }

        private static string SnapshotWithoutStealPath(WdSession s, string openPath)
        {
            string snap = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-ooxml-snap.docx");
            string saveHow = "";
            bool saved = false;
            s.TryWpsDoc(openPath, doc =>
            {
                try
                {
                    Invoke(doc, "Save");
                    saved = true;
                    saveHow = "Save";
                    return true;
                }
                catch (Exception ex)
                {
                    saveHow = "Save 失败: " + TrimErr(ex.Message);
                }

                try
                {
                    Invoke(doc, "SaveAs", snap);
                    saveHow = saveHow + " | SaveAs副本";
                    saved = File.Exists(snap);
                }
                catch (Exception ex)
                {
                    saveHow = saveHow + " | SaveAs 失败: " + TrimErr(ex.Message);
                }

                try
                {
                    Invoke(doc, "SaveAs", openPath);
                    saveHow = saveHow + " | SaveAs回原路径";
                }
                catch (Exception ex)
                {
                    saveHow = saveHow + " | 回写失败: " + TrimErr(ex.Message);
                }

                return saved;
            });
            Line("落盘", saveHow + (saved ? " 成功" : " 失败"));

            if (File.Exists(openPath))
            {
                try
                {
                    File.Copy(openPath, snap, true);
                    return snap;
                }
                catch (Exception ex)
                {
                    Line("复制打开文件", TrimErr(ex.Message));
                }
            }

            return File.Exists(snap) ? snap : null;
        }

        private static string ReadChartSeriesFromWordOpenXml(object doc)
        {
            string raw;
            try
            {
                raw = Convert.ToString(GetProp(GetProp(doc, "Content"), "WordOpenXML") ?? "");
            }
            catch (Exception ex)
            {
                return "取 WordOpenXML 失败: " + TrimErr(ex.Message);
            }

            if (string.IsNullOrEmpty(raw))
            {
                return "WordOpenXML 空";
            }

            try
            {
                var xml = new XmlDocument();
                xml.LoadXml(raw);
                XmlNodeList parts = xml.SelectNodes("//*[local-name()='part']");
                var names = new List<string>();
                if (parts != null)
                {
                    foreach (XmlNode part in parts)
                    {
                        string name = "";
                        if (part.Attributes != null)
                        {
                            foreach (XmlAttribute a in part.Attributes)
                            {
                                if (a.LocalName == "name")
                                {
                                    name = a.Value ?? "";
                                    break;
                                }
                            }
                        }

                        if (name.IndexOf("chart", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            names.Add(name);
                        }
                    }
                }

                var sers = xml.SelectNodes("//*[local-name()='ser']");
                var sb = new StringBuilder();
                sb.Append("len=").Append(raw.Length)
                    .Append(" chart部件=").Append(names.Count == 0 ? "无" : string.Join("|", names))
                    .Append(" ser=").Append(sers == null ? 0 : sers.Count);
                if (sers != null)
                {
                    int i = 0;
                    foreach (XmlNode ser in sers)
                    {
                        i++;
                        sb.Append(" #").Append(i)
                            .Append(" name=").Append(FirstLocalValue(ser, "tx", "v"))
                            .Append(" values=[").Append(JoinLocalPts(ser, "val")).Append(']');
                    }
                }

                return sb.ToString();
            }
            catch (Exception ex)
            {
                return "解析失败: " + TrimErr(ex.Message);
            }
        }

        private static string DumpEmbedFromWordOpenXml(object doc)
        {
            string raw;
            try
            {
                raw = Convert.ToString(GetProp(GetProp(doc, "Content"), "WordOpenXML") ?? "");
            }
            catch (Exception ex)
            {
                return "取 WordOpenXML 失败: " + TrimErr(ex.Message);
            }

            if (string.IsNullOrEmpty(raw))
            {
                return "空";
            }

            try
            {
                var xml = new XmlDocument();
                xml.LoadXml(raw);
                XmlNodeList parts = xml.SelectNodes("//*[local-name()='part']");
                var names = new List<string>();
                int binary = 0;
                if (parts != null)
                {
                    foreach (XmlNode part in parts)
                    {
                        string name = PartName(part);
                        if (name.IndexOf("embed", StringComparison.OrdinalIgnoreCase) >= 0
                            || name.IndexOf(".xlsx", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            int xmlLen = part.InnerText == null ? 0 : part.InnerText.Length;
                            names.Add(name + " inner=" + xmlLen);
                            if (name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                            {
                                binary++;
                            }
                        }
                    }
                }

                string grid = "";
                if (parts != null)
                {
                    foreach (XmlNode part in parts)
                    {
                        string name = PartName(part);
                        if (!name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        byte[] xlsx = TryPartBinary(part);
                        if (xlsx != null && xlsx.Length > 0)
                        {
                            grid = ReadSheetGridFromXlsx(xlsx);
                        }

                        break;
                    }
                }

                return names.Count == 0
                    ? "无 embeddings 部件"
                    : string.Join(" | ", names) + " xlsx部件=" + binary + " " + grid;
            }
            catch (Exception ex)
            {
                return "解析失败: " + TrimErr(ex.Message);
            }
        }

        private static string PartName(XmlNode part)
        {
            if (part == null || part.Attributes == null)
            {
                return "";
            }

            foreach (XmlAttribute a in part.Attributes)
            {
                if (a.LocalName == "name")
                {
                    return a.Value ?? "";
                }
            }

            return "";
        }

        private static byte[] TryPartBinary(XmlNode part)
        {
            if (part == null)
            {
                return null;
            }

            XmlNode data = null;
            XmlNodeList nodes = part.SelectNodes(".//*[local-name()='binaryData']");
            if (nodes != null && nodes.Count > 0)
            {
                data = nodes[0];
            }

            if (data == null || string.IsNullOrEmpty(data.InnerText))
            {
                return null;
            }

            try
            {
                string b64 = data.InnerText.Replace("\r", "").Replace("\n", "").Replace(" ", "");
                return Convert.FromBase64String(b64);
            }
            catch
            {
                return null;
            }
        }

        private static string ReadEmbeddedSheetFromDocx(string path)
        {
            if (!File.Exists(path))
            {
                return "文件不存在";
            }

            try
            {
                using (ZipArchive z = ZipFile.OpenRead(path))
                {
                    foreach (ZipArchiveEntry e in z.Entries)
                    {
                        string n = e.FullName.Replace('\\', '/');
                        if (n.IndexOf("word/embeddings/", StringComparison.OrdinalIgnoreCase) < 0
                            || !n.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        byte[] raw;
                        using (Stream input = e.Open())
                        using (var ms = new MemoryStream())
                        {
                            input.CopyTo(ms);
                            raw = ms.ToArray();
                        }

                        return n + " " + raw.Length + "B " + ReadSheetGridFromXlsx(raw);
                    }
                }
            }
            catch (Exception ex)
            {
                return "解包失败: " + TrimErr(ex.Message);
            }

            return "无 xlsx 嵌入";
        }

        private static string ReadSheetGridFromXlsx(byte[] xlsx)
        {
            var shared = new List<string>();
            var cells = new List<string>();
            try
            {
                using (var input = new MemoryStream(xlsx))
                using (var z = new ZipArchive(input, ZipArchiveMode.Read))
                {
                    ZipArchiveEntry sst = null;
                    var sheets = new List<ZipArchiveEntry>();
                    foreach (ZipArchiveEntry e in z.Entries)
                    {
                        string n = e.FullName.Replace('\\', '/');
                        if (n.IndexOf("sharedStrings", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            sst = e;
                        }
                        else if (n.IndexOf("sheet", StringComparison.OrdinalIgnoreCase) >= 0
                            && n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                        {
                            sheets.Add(e);
                        }
                    }

                    if (sst != null)
                    {
                        var xml = new XmlDocument();
                        using (Stream s = sst.Open())
                        {
                            xml.Load(s);
                        }

                        XmlNodeList si = xml.SelectNodes("//*[local-name()='t']");
                        if (si != null)
                        {
                            foreach (XmlNode t in si)
                            {
                                shared.Add((t.InnerText ?? "").Trim());
                            }
                        }
                    }

                    foreach (ZipArchiveEntry sheet in sheets)
                    {
                        var xml = new XmlDocument();
                        using (Stream s = sheet.Open())
                        {
                            xml.Load(s);
                        }

                        XmlNodeList cs = xml.SelectNodes("//*[local-name()='c']");
                        if (cs == null)
                        {
                            continue;
                        }

                        foreach (XmlNode c in cs)
                        {
                            string r = "";
                            string t = "";
                            if (c.Attributes != null)
                            {
                                foreach (XmlAttribute a in c.Attributes)
                                {
                                    if (a.LocalName == "r")
                                    {
                                        r = a.Value ?? "";
                                    }
                                    else if (a.LocalName == "t")
                                    {
                                        t = a.Value ?? "";
                                    }
                                }
                            }

                            string v = "";
                            foreach (XmlNode child in c.ChildNodes)
                            {
                                if (child.LocalName == "v")
                                {
                                    v = (child.InnerText ?? "").Trim();
                                    break;
                                }
                            }

                            if (t == "s")
                            {
                                int i;
                                if (int.TryParse(v, out i) && i >= 0 && i < shared.Count)
                                {
                                    v = shared[i];
                                }
                            }

                            if (!string.IsNullOrEmpty(r) || !string.IsNullOrEmpty(v))
                            {
                                cells.Add(r + "=" + v);
                            }

                            if (cells.Count >= 24)
                            {
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return "读表失败: " + TrimErr(ex.Message);
            }

            return "shared=[" + string.Join(",", shared) + "] cells=[" + string.Join(" ", cells) + "]";
        }

        private static string ReadChartSeriesFromDocx(string path)
        {
            if (!File.Exists(path))
            {
                return "文件不存在";
            }

            var sb = new StringBuilder();
            try
            {
                using (ZipArchive z = ZipFile.OpenRead(path))
                {
                    foreach (ZipArchiveEntry e in z.Entries)
                    {
                        string n = e.FullName.Replace('\\', '/');
                        if (n.IndexOf("word/charts/chart", StringComparison.OrdinalIgnoreCase) < 0
                            || !n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                            || n.IndexOf("colors", StringComparison.OrdinalIgnoreCase) >= 0
                            || n.IndexOf("style", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            continue;
                        }

                        var xml = new XmlDocument();
                        using (Stream input = e.Open())
                        {
                            xml.Load(input);
                        }

                        var sers = xml.SelectNodes("//*[local-name()='ser']");
                        sb.Append(n).Append(" ser=").Append(sers == null ? 0 : sers.Count);
                        if (sers == null)
                        {
                            continue;
                        }

                        int i = 0;
                        foreach (XmlNode ser in sers)
                        {
                            i++;
                            string name = FirstLocalValue(ser, "tx", "v");
                            string cats = JoinLocalPts(ser, "cat");
                            string vals = JoinLocalPts(ser, "val");
                            sb.Append(" #").Append(i)
                                .Append(" name=").Append(name)
                                .Append(" cats=[").Append(cats).Append(']')
                                .Append(" values=[").Append(vals).Append(']');
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return "解包失败: " + TrimErr(ex.Message);
            }

            return sb.Length == 0 ? "无 chart xml" : sb.ToString();
        }

        private static string FirstLocalValue(XmlNode parent, string scope, string local)
        {
            XmlNode scopeNode = null;
            foreach (XmlNode child in parent.ChildNodes)
            {
                if (child.LocalName == scope)
                {
                    scopeNode = child;
                    break;
                }
            }

            if (scopeNode == null)
            {
                return "";
            }

            XmlNodeList vs = scopeNode.SelectNodes(".//*[local-name()='" + local + "']");
            return vs == null || vs.Count == 0 ? "" : (vs[0].InnerText ?? "").Trim();
        }

        private static string JoinLocalPts(XmlNode ser, string kind)
        {
            XmlNode parent = null;
            foreach (XmlNode child in ser.ChildNodes)
            {
                if (child.LocalName == kind)
                {
                    parent = child;
                    break;
                }
            }

            if (parent == null)
            {
                return "";
            }

            var parts = new List<string>();
            XmlNodeList pts = parent.SelectNodes(".//*[local-name()='pt']");
            if (pts == null)
            {
                return "";
            }

            foreach (XmlNode pt in pts)
            {
                XmlNode v = null;
                foreach (XmlNode child in pt.ChildNodes)
                {
                    if (child.LocalName == "v")
                    {
                        v = child;
                        break;
                    }
                }

                parts.Add(v == null ? "" : (v.InnerText ?? ""));
                if (parts.Count >= 8)
                {
                    break;
                }
            }

            return string.Join(",", parts);
        }

        private static int DeleteInlineShapes(object doc)
        {
            object inlines = GetProp(doc, "InlineShapes");
            int n = Convert.ToInt32(GetProp(inlines, "Count"));
            int deleted = 0;
            for (int i = n; i >= 1; i--)
            {
                try
                {
                    object shape = Invoke(inlines, "Item", i);
                    Invoke(shape, "Delete");
                    deleted++;
                }
                catch (Exception ex)
                {
                    Line("删图#" + i, TrimErr(ex.Message));
                }
            }

            return deleted;
        }

        private static bool TryClickWakePoints(WdSession s, string path)
        {
            var screen = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
            Line("主屏", screen.Left + "," + screen.Top + " " + screen.Width + "x" + screen.Height);
            DumpTopWindows("before");
            IntPtr pane = IntPtr.Zero;
            IntPtr root = IntPtr.Zero;
            int gl = 0, gt = 0, gw = 0, gh = 0;
            bool havePoint = false;
            s.TryWpsDoc(path, doc =>
            {
                try
                {
                    Invoke(doc, "Activate");
                }
                catch
                {
                }

                object app = GetProp(doc, "Application");
                ActivateWindowForDoc(app, path);
                object window = TryWindow(doc, app);
                object shape = FirstInline(doc);
                if (shape == null)
                {
                    Line("点击", "无 InlineShape");
                    return false;
                }

                object range = GetProp(shape, "Range");
                try
                {
                    Invoke(range, "Select");
                }
                catch
                {
                }

                if (window != null)
                {
                    try
                    {
                        Invoke(window, "ScrollIntoView", range);
                    }
                    catch
                    {
                    }
                }

                IntPtr hwnd = TryHwnd(window);
                if (hwnd == IntPtr.Zero)
                {
                    hwnd = TryHwnd(app);
                }

                root = hwnd == IntPtr.Zero ? IntPtr.Zero : GetAncestor(hwnd, GaRoot);
                if (root == IntPtr.Zero)
                {
                    root = hwnd;
                }

                Line("docHwnd", hwnd.ToInt64() + " root=" + root.ToInt64() + " iconic=" + (root != IntPtr.Zero && IsIconic(root)));
                RestoreTop(root);
                BringWindowFront(root);
                System.Threading.Thread.Sleep(200);
                pane = FindClass(root != IntPtr.Zero ? root : hwnd, "_WwG");
                if (pane == IntPtr.Zero)
                {
                    pane = hwnd;
                }

                if (GetWindowRect(root, out RECT rr))
                {
                    Line("顶层矩形", rr.Left + "," + rr.Top + " " + (rr.Right - rr.Left) + "x" + (rr.Bottom - rr.Top));
                }

                if (pane != IntPtr.Zero && GetWindowRect(pane, out RECT pr))
                {
                    Line("_WwG", pane.ToInt64() + " " + pr.Left + "," + pr.Top + " " + (pr.Right - pr.Left) + "x" + (pr.Bottom - pr.Top));
                }

                havePoint = window != null && TryGetPoint(window, range, out gl, out gt, out gw, out gh);
                Line("GetPoint", havePoint ? gl + "," + gt + " " + gw + "x" + gh : "失败");
                DumpTopWindows("after-activate");
                return true;
            });

            if (havePoint && pane != IntPtr.Zero)
            {
                var screenPt = new POINT { X = gl + Math.Max(12, gw / 2), Y = gt + Math.Max(12, gh / 6) };
                var client = screenPt;
                ScreenToClient(pane, ref client);
                Line("发消息", "screen=" + screenPt.X + "," + screenPt.Y + " client=" + client.X + "," + client.Y);
                RestoreTop(root);
                BringWindowFront(root);
                SendClickToWindow(pane, client.X, client.Y);
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(800);
                string afterMsg = DescribeCharts(s, path);
                Line("消息点后", afterMsg);
                if (LooksLiveDump(afterMsg))
                {
                    return true;
                }
            }

            var clicks = new List<Tuple<string, int, int>>();
            if (havePoint)
            {
                AddClick(clicks, "图上沿", gl + gw / 2, gt + Math.Min(60, Math.Max(12, gh / 6)));
                AddClick(clicks, "图中心", gl + gw / 2, gt + gh / 2);
                var vis = IntersectScreen(gl, gt, gw, gh, screen);
                if (vis != null)
                {
                    AddClick(clicks, "可见区中心", vis.Item1, vis.Item2);
                }
            }

            if (pane != IntPtr.Zero && GetWindowRect(pane, out RECT paneRect) && paneRect.Left > -10000)
            {
                AddClick(clicks, "文档区中上", paneRect.Left + (paneRect.Right - paneRect.Left) / 2, paneRect.Top + 120);
            }

            foreach (Tuple<string, int, int> p in clicks)
            {
                if (p.Item2 < -10000 || p.Item3 < -10000)
                {
                    Line("跳过 " + p.Item1, p.Item2 + "," + p.Item3 + " 最小化坐标");
                    continue;
                }

                int x = Clamp(p.Item2, screen.Left + 8, screen.Right - 8);
                int y = Clamp(p.Item3, screen.Top + 8, screen.Bottom - 8);
                Line("真鼠标 " + p.Item1, p.Item2 + "," + p.Item3 + (x != p.Item2 || y != p.Item3 ? " 夹到 " + x + "," + y : ""));
                RestoreTop(root);
                BringWindowFront(root);
                ClickScreen(x, y);
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(700);
                string after = DescribeCharts(s, path);
                Line("点后", after);
                if (LooksLiveDump(after))
                {
                    return true;
                }
            }

            if (TryOleDoVerb(s, path))
            {
                return true;
            }

            return false;
        }

        private static bool TryOleDoVerb(WdSession s, string path)
        {
            s.TryWpsDoc(path, doc =>
            {
                object shape = FirstInline(doc);
                if (shape == null)
                {
                    Line("DoVerb", "无图");
                    return false;
                }

                try
                {
                    Invoke(shape, "Select");
                }
                catch
                {
                }

                object ole = null;
                try
                {
                    ole = GetProp(shape, "OLEFormat");
                    Line("OLEFormat", ole == null ? "null" : ole.GetType().Name);
                }
                catch (Exception ex)
                {
                    Line("OLEFormat", TrimErr(ex.Message));
                }

                if (ole != null)
                {
                    foreach (int verb in new[] { 0, -1, -2, 1 })
                    {
                        try
                        {
                            Invoke(ole, "DoVerb", verb);
                            Line("DoVerb " + verb, "成功");
                        }
                        catch (Exception ex)
                        {
                            Line("DoVerb " + verb, TrimErr(ex.Message));
                        }
                    }

                    try
                    {
                        Invoke(ole, "Activate");
                        Line("OLE.Activate", "成功");
                    }
                    catch (Exception ex)
                    {
                        Line("OLE.Activate", TrimErr(ex.Message));
                    }
                }

                object chart = TryChart(shape);
                if (chart != null)
                {
                    try
                    {
                        Invoke(chart, "Activate");
                        Line("Chart.Activate", "成功");
                    }
                    catch (Exception ex)
                    {
                        Line("Chart.Activate", TrimErr(ex.Message));
                    }
                }

                return true;
            });
            System.Windows.Forms.Application.DoEvents();
            System.Threading.Thread.Sleep(800);
            string after = DescribeCharts(s, path);
            Line("DoVerb后", after);
            return LooksLiveDump(after);
        }

        private static void RestoreTop(IntPtr root)
        {
            if (root == IntPtr.Zero)
            {
                return;
            }

            try
            {
                if (IsIconic(root) || !IsWindowVisible(root))
                {
                    ShowWindow(root, SwRestore);
                }
            }
            catch
            {
            }
        }

        private static void SendClickToWindow(IntPtr hwnd, int clientX, int clientY)
        {
            IntPtr lp = (IntPtr)((clientY << 16) | (clientX & 0xFFFF));
            IntPtr mk = (IntPtr)1;
            SendMessage(hwnd, WmMousemove, IntPtr.Zero, lp);
            SendMessage(hwnd, WmLbuttondown, mk, lp);
            System.Threading.Thread.Sleep(40);
            SendMessage(hwnd, WmLbuttonup, IntPtr.Zero, lp);
        }

        private static IntPtr FindClass(IntPtr root, string className)
        {
            IntPtr found = IntPtr.Zero;
            if (root == IntPtr.Zero)
            {
                return IntPtr.Zero;
            }

            EnumChildProc cb = (child, _) =>
            {
                var sb = new StringBuilder(64);
                GetClassName(child, sb, sb.Capacity);
                if (string.Equals(sb.ToString(), className, StringComparison.Ordinal))
                {
                    if (GetWindowRect(child, out RECT r) && r.Left > -10000 && (r.Right - r.Left) > 80)
                    {
                        found = child;
                        return false;
                    }

                    if (found == IntPtr.Zero)
                    {
                        found = child;
                    }
                }

                return true;
            };
            EnumChildWindows(root, cb, IntPtr.Zero);
            return found;
        }

        private static void DumpTopWindows(string hint)
        {
            var names = new List<string>();
            EnumWindowsProc cb = (h, _) =>
            {
                if (!IsWindowVisible(h))
                {
                    return true;
                }

                var title = new StringBuilder(160);
                GetWindowText(h, title, title.Capacity);
                string t = title.ToString();
                if (t.IndexOf("WPS", StringComparison.OrdinalIgnoreCase) < 0
                    && t.IndexOf("yiwrite", StringComparison.OrdinalIgnoreCase) < 0
                    && t.IndexOf(".docx", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return true;
                }

                var cls = new StringBuilder(64);
                GetClassName(h, cls, cls.Capacity);
                GetWindowRect(h, out RECT r);
                names.Add(cls + " '" + t + "' " + r.Left + "," + r.Top + " " + (r.Right - r.Left) + "x" + (r.Bottom - r.Top));
                return names.Count < 10;
            };
            EnumWindows(cb, IntPtr.Zero);
            Line("顶层含WPS/" + hint, names.Count == 0 ? "无" : string.Join(" || ", names));
        }

        private static void AddClick(List<Tuple<string, int, int>> points, string name, int x, int y)
        {
            points.Add(Tuple.Create(name, x, y));
        }

        private static Tuple<int, int> IntersectScreen(int left, int top, int width, int height, System.Drawing.Rectangle screen)
        {
            int r = left + width;
            int b = top + height;
            int il = Math.Max(left, screen.Left);
            int it = Math.Max(top, screen.Top);
            int ir = Math.Min(r, screen.Right);
            int ib = Math.Min(b, screen.Bottom);
            if (ir - il < 20 || ib - it < 20)
            {
                return null;
            }

            return Tuple.Create((il + ir) / 2, (it + ib) / 2);
        }

        private static int Clamp(int v, int min, int max)
        {
            if (v < min)
            {
                return min;
            }

            if (v > max)
            {
                return max;
            }

            return v;
        }

        private static void ActivateWindowForDoc(object app, string path)
        {
            string name = Path.GetFileName(path);
            try
            {
                Line("ActiveDoc", Convert.ToString(GetProp(GetProp(app, "ActiveDocument"), "FullName")));
            }
            catch (Exception ex)
            {
                Line("ActiveDoc", TrimErr(ex.Message));
            }

            try
            {
                object windows = GetProp(app, "Windows");
                int n = Convert.ToInt32(GetProp(windows, "Count"));
                Line("Windows", n.ToString());
                for (int i = 1; i <= n; i++)
                {
                    object w = Invoke(windows, "Item", i);
                    string cap = "";
                    try
                    {
                        cap = Convert.ToString(GetProp(w, "Caption"));
                    }
                    catch
                    {
                    }

                    Line("Window#" + i, cap);
                    if (!string.IsNullOrEmpty(name)
                        && cap != null
                        && cap.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        try
                        {
                            Invoke(w, "Activate");
                            Line("切窗口", cap);
                        }
                        catch (Exception ex)
                        {
                            Line("切窗口", TrimErr(ex.Message));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Line("Windows", TrimErr(ex.Message));
            }

            IntPtr byTitle = FindTopByTitle(name);
            Line("按标题找窗", byTitle == IntPtr.Zero ? "无" : byTitle.ToInt64().ToString());
            if (byTitle != IntPtr.Zero)
            {
                RestoreTop(byTitle);
                BringWindowFront(byTitle);
            }
        }

        private static IntPtr FindTopByTitle(string hint)
        {
            IntPtr found = IntPtr.Zero;
            if (string.IsNullOrEmpty(hint))
            {
                return IntPtr.Zero;
            }

            EnumWindowsProc cb = (h, _) =>
            {
                var title = new StringBuilder(200);
                GetWindowText(h, title, title.Capacity);
                string t = title.ToString();
                if (t.IndexOf(hint, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    found = h;
                    return false;
                }

                return true;
            };
            EnumWindows(cb, IntPtr.Zero);
            return found;
        }

        private static object TryWindow(object doc, object app)
        {
            try
            {
                object w = GetProp(doc, "ActiveWindow");
                if (w != null)
                {
                    return w;
                }
            }
            catch
            {
            }

            try
            {
                return GetProp(app, "ActiveWindow");
            }
            catch
            {
                return null;
            }
        }

        private static void BringDocFront(WdSession s, string path)
        {
            s.TryWpsDoc(path, doc =>
            {
                try
                {
                    Invoke(doc, "Activate");
                    object app = GetProp(doc, "Application");
                    object window = TryWindow(doc, app);
                    try
                    {
                        Invoke(app, "Activate");
                    }
                    catch
                    {
                    }

                    BringWindowFront(TryHwnd(window) != IntPtr.Zero ? TryHwnd(window) : TryHwnd(app));
                }
                catch
                {
                }

                return true;
            });
        }

        private static void DumpChildWindows(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
            {
                return;
            }

            var names = new List<string>();
            EnumChildWindows(hwnd, (child, _) =>
            {
                var sb = new StringBuilder(64);
                GetClassName(child, sb, sb.Capacity);
                if (GetWindowRect(child, out RECT r))
                {
                    int w = r.Right - r.Left;
                    int h = r.Bottom - r.Top;
                    if (w > 80 && h > 80 && names.Count < 8)
                    {
                        names.Add(sb + " " + w + "x" + h + "@" + r.Left + "," + r.Top);
                    }
                }

                return true;
            }, IntPtr.Zero);
            Line("子窗", names.Count == 0 ? "无" : string.Join(" | ", names));
        }

        private static bool TryGetPoint(object window, object range, out int left, out int top, out int width, out int height)
        {
            left = top = width = height = 0;
            try
            {
                object[] args = { 0, 0, 0, 0, range };
                var mods = new ParameterModifier(5);
                mods[0] = true;
                mods[1] = true;
                mods[2] = true;
                mods[3] = true;
                window.GetType().InvokeMember(
                    "GetPoint",
                    BindingFlags.InvokeMethod,
                    null,
                    window,
                    args,
                    new ParameterModifier[] { mods },
                    null,
                    null);
                left = Convert.ToInt32(args[0]);
                top = Convert.ToInt32(args[1]);
                width = Convert.ToInt32(args[2]);
                height = Convert.ToInt32(args[3]);
                return width > 0 && height > 0;
            }
            catch (Exception ex)
            {
                Line("GetPoint", TrimErr(ex.Message));
                return false;
            }
        }

        private static IntPtr TryHwnd(object target)
        {
            if (target == null)
            {
                return IntPtr.Zero;
            }

            try
            {
                object h = GetProp(target, "Hwnd");
                if (h == null)
                {
                    return IntPtr.Zero;
                }

                return new IntPtr(Convert.ToInt64(h));
            }
            catch
            {
                return IntPtr.Zero;
            }
        }

        private static void BringWindowFront(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
            {
                return;
            }

            try
            {
                IntPtr root = GetAncestor(hwnd, GaRoot);
                if (root == IntPtr.Zero)
                {
                    root = hwnd;
                }

                RestoreTop(root);
                keybd_event(VkMenu, 0, 0, UIntPtr.Zero);
                keybd_event(VkMenu, 0, KeyeventfKeyup, UIntPtr.Zero);
                IntPtr fore = GetForegroundWindow();
                uint foreTid = GetWindowThreadProcessId(fore, out _);
                uint selfTid = GetCurrentThreadId();
                if (foreTid != 0 && foreTid != selfTid)
                {
                    AttachThreadInput(selfTid, foreTid, true);
                    SetForegroundWindow(root);
                    AttachThreadInput(selfTid, foreTid, false);
                }
                else
                {
                    SetForegroundWindow(root);
                }
            }
            catch
            {
                try
                {
                    SetForegroundWindow(hwnd);
                }
                catch
                {
                }
            }
        }

        private static void ClickScreen(int x, int y)
        {
            SetCursorPos(x, y);
            System.Threading.Thread.Sleep(80);
            mouse_event(MouseeventfLeftdown, 0, 0, 0, UIntPtr.Zero);
            System.Threading.Thread.Sleep(40);
            mouse_event(MouseeventfLeftup, 0, 0, 0, UIntPtr.Zero);
        }

        private static void TryWake(WdSession s, string path, string tag)
        {
            s.TryWpsDoc(path, doc =>
            {
                object inlines = GetProp(doc, "InlineShapes");
                int n = Convert.ToInt32(GetProp(inlines, "Count"));
                if (n < 1)
                {
                    Line(tag, "无图");
                    return true;
                }

                object shape = Invoke(inlines, "Item", 1);
                try
                {
                    Invoke(GetProp(shape, "Range"), "Select");
                }
                catch (Exception ex)
                {
                    Line(tag + " Select", TrimErr(ex.Message));
                }

                object chart = TryChart(shape);
                if (chart == null)
                {
                    Line(tag, "无 Chart 对象");
                    return true;
                }

                foreach (string name in new[] { "Activate", "Refresh", "Select" })
                {
                    try
                    {
                        Invoke(chart, name);
                        Line(tag + " " + name, "成功");
                    }
                    catch (Exception ex)
                    {
                        Line(tag + " " + name, TrimErr(ex.Message));
                    }
                }

                try
                {
                    object app = GetProp(doc, "Application");
                    Invoke(app, "ScreenRefresh");
                    Line(tag + " ScreenRefresh", "成功");
                }
                catch (Exception ex)
                {
                    Line(tag + " ScreenRefresh", TrimErr(ex.Message));
                }

                return true;
            });
        }

        public static int RunAutoCopyToBlank()
        {
            var s = new WdSession("wps");
            string src = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-src.docx");
            string blank = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-blank.docx");
            if (!File.Exists(src) || !File.Exists(blank))
            {
                Console.WriteLine("找不到样例稿");
                return 1;
            }

            ToolResult openSrc = s.Open(src, false, "wps");
            ToolResult openBlank = s.Open(blank, false, "wps");
            if (!openSrc.Success || !openBlank.Success)
            {
                Console.WriteLine("打开失败 src=" + (openSrc.Error ?? "") + " blank=" + (openBlank.Error ?? ""));
                return 1;
            }

            bool copied = false;
            string err = "";
            s.TryWpsDoc(src, srcDoc =>
            {
                object shape = FirstInline(srcDoc);
                if (shape == null)
                {
                    err = "带图稿没有 InlineShape";
                    return false;
                }

                try
                {
                    TrySilenceAlerts(srcDoc);
                    object range = GetProp(shape, "Range");
                    Invoke(range, "Select");
                    Invoke(range, "Copy");
                }
                catch (Exception ex)
                {
                    err = "Copy 失败: " + TrimErr(ex.Message);
                    return false;
                }

                return s.TryWpsDoc(blank, destDoc =>
                {
                    TrySilenceAlerts(destDoc);
                    object content = GetProp(destDoc, "Content");
                    int end = Convert.ToInt32(GetProp(content, "End"));
                    int pos = end > 1 ? end - 1 : end;
                    object at = Invoke(destDoc, "Range", pos, pos);
                    string last = "";
                    if (TryCall("Paste(content)", () => Invoke(content, "Paste")))
                    {
                        copied = true;
                        return true;
                    }

                    if (TryCall("Paste(endRange)", () => Invoke(at, "Paste")))
                    {
                        copied = true;
                        return true;
                    }

                    try
                    {
                        Invoke(at, "Select");
                        object sel = GetProp(GetProp(destDoc, "Application"), "Selection");
                        Invoke(sel, "Paste");
                        last = "Paste(selection)";
                        copied = true;
                        return true;
                    }
                    catch (Exception ex)
                    {
                        err = last + " " + TrimErr(ex.Message);
                        return false;
                    }
                });
            });

            Console.WriteLine(copied ? "已用 COM Select+Copy+Paste 贴到空稿末尾。" : ("失败: " + err));
            Console.WriteLine("形状: " + DescribeCharts(s, blank));
            Console.WriteLine("请点刚贴进去的那张，看能不能改数。窗口不关。");
            s.ActivateDoc(blank);
            return copied ? 0 : 1;
        }

        public static int RunInspectManualDest()
        {
            var s = new WdSession("wps");
            string blank = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-blank.docx");
            if (!File.Exists(blank))
            {
                Console.WriteLine("找不到空稿: " + blank);
                return 1;
            }

            ToolResult opened = s.Open(blank, false, "wps");
            if (!opened.Success)
            {
                Console.WriteLine("打开空稿失败: " + (opened.Error ?? ""));
                return 1;
            }

            string dump = DescribeCharts(s, blank);
            Console.WriteLine("形状: " + dump);
            Console.WriteLine("活图: " + (LooksLiveDump(dump) ? "是" : "否"));
            ToolResult got = s.GetContent(WdSession.Str(opened, "channel_id"));
            string body = ReadBody(got);
            Console.WriteLine("读文 <Chart: " + ((body ?? "").IndexOf("<Chart", StringComparison.OrdinalIgnoreCase) >= 0 ? "是" : "否"));
            Console.WriteLine("读文片段: " + Snip(body, 320));
            string fromOpenXml = "";
            s.TryWpsDoc(blank, doc =>
            {
                fromOpenXml = ReadChartSeriesFromWordOpenXml(doc);
                return true;
            });
            Console.WriteLine("WordOpenXML 套数: " + fromOpenXml);
            string embedFromOpenXml = "";
            s.TryWpsDoc(blank, doc =>
            {
                embedFromOpenXml = DumpEmbedFromWordOpenXml(doc);
                return true;
            });
            Console.WriteLine("WordOpenXML 内嵌表: " + embedFromOpenXml);
            string snap = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-ooxml-snap.docx");
            if (File.Exists(snap))
            {
                Console.WriteLine("磁盘副本内嵌表: " + ReadEmbeddedSheetFromDocx(snap));
            }

            ProbeWriteSeries(s, blank);
            return LooksLiveDump(dump) ? 0 : 1;
        }

        public static int RunManualSamples()
        {
            var s = new WdSession("wps");
            string src = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-src.docx");
            string blank = Path.Combine(Path.GetTempPath(), "yiwrite-wps-chart-blank.docx");
            string seed = s.CreateChartFixtureDocx(1);
            if (string.IsNullOrEmpty(seed) || !File.Exists(seed))
            {
                Console.WriteLine("FAIL Word 做带图种子失败");
                return 1;
            }

            File.Copy(seed, src, true);
            ToolResult openSrc = s.Open(src, false, "wps");
            if (!openSrc.Success)
            {
                Console.WriteLine("FAIL 打开带图稿: " + (openSrc.Error ?? ""));
                return 1;
            }

            ToolResult openBlank = s.Open(blank, true, "wps");
            if (!openBlank.Success)
            {
                Console.WriteLine("FAIL 打开空稿: " + (openBlank.Error ?? ""));
                return 1;
            }

            s.TryWpsDoc(blank, doc =>
            {
                try
                {
                    object content = GetProp(doc, "Content");
                    SetProp(content, "Text", "这是空稿，请把左边那张图复制粘贴到这里。\r");
                    return true;
                }
                catch
                {
                    return false;
                }
            });
            s.ActivateDoc(src);
            s.ActivateDoc(blank);
            s.ActivateDoc(src);

            Console.WriteLine();
            Console.WriteLine("两份都已在 WPS 打开，不会关。");
            Console.WriteLine("带图: " + src);
            Console.WriteLine("空稿: " + blank);
            Console.WriteLine("请在带图稿里点选图表，Ctrl+C，切到空稿，Ctrl+V。");
            Console.WriteLine("看粘贴后能不能再点开图表改数。完了告诉我结果。");
            return 0;
        }

        public static int RunWhyDead()
        {
            var s = new WdSession("wps");
            WdSession.SweepBeforeRun("wps");
            try
            {
                string seedRaw = s.CreateChartFixtureDocx(1);
                if (string.IsNullOrEmpty(seedRaw) || !File.Exists(seedRaw))
                {
                    Console.WriteLine("FAIL 种子失败");
                    return 1;
                }

                string seed = Path.Combine(
                    Path.GetTempPath(),
                    "ew-wd-seed-keep-" + Guid.NewGuid().ToString("N") + ".docx");
                File.Copy(seedRaw, seed, true);
                Line("种子", seed);

                Section("A 直接打开后另存", () =>
                {
                    ProbeOpenExisting(s, seed, "打开");
                    string snap = SaveOpenDocCopy(s, seed, "打开后另存");
                    return snap == null ? 1 : DumpChartParts(snap, "打开后另存");
                });

                string insertPath = s.NewDocxPath();
                string insertId = null;
                Section("B InsertFile 后另存", () =>
                {
                    int n = ProbeInsertFile(s, seed, insertPath, out insertId, comInsert: true);
                    DumpOpenXmlHints(s, insertPath, "InsertFile 现场 WordOpenXML");
                    string snap = SaveOpenDocCopy(s, insertPath, "InsertFile后另存");
                    if (snap != null)
                    {
                        DumpChartParts(snap, "InsertFile后另存");
                    }

                    return n;
                });

                Section("C PasteSpecial 几档", () => ProbePasteFlavors(s, seed));
                return 0;
            }
            finally
            {
                WdSession.SweepAfterRun("wps");
            }
        }

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

        private static string SaveOpenDocCopy(WdSession s, string openPath, string tag)
        {
            string dest = Path.Combine(
                Path.GetTempPath(),
                "ew-wd-snap-" + Guid.NewGuid().ToString("N") + ".docx");
            bool saved = s.TryWpsDoc(openPath, doc =>
            {
                try
                {
                    Invoke(doc, "SaveAs", dest);
                    return File.Exists(dest);
                }
                catch (Exception ex)
                {
                    Line(tag + " SaveAs", TrimErr(ex.Message));
                    return false;
                }
            });
            if (!saved)
            {
                Line(tag, "另存失败");
                return null;
            }

            string copy = dest + ".read.docx";
            try
            {
                File.Copy(dest, copy, true);
                return copy;
            }
            catch
            {
                return dest;
            }
        }

        private static void DumpOpenXmlHints(WdSession s, string fullPath, string tag)
        {
            string xml = "";
            s.TryWpsDoc(fullPath, doc =>
            {
                try
                {
                    object content = GetProp(doc, "Content");
                    xml = Convert.ToString(GetProp(content, "WordOpenXML") ?? "");
                    return true;
                }
                catch (Exception ex)
                {
                    Line(tag, "取 WordOpenXML 失败: " + TrimErr(ex.Message));
                    return false;
                }
            });
            if (string.IsNullOrEmpty(xml))
            {
                Line(tag, "空");
                return;
            }

            Line(tag + " 长度", xml.Length.ToString());
            Line(tag + " 含 c:chart", xml.IndexOf("chart", StringComparison.OrdinalIgnoreCase) >= 0 ? "是" : "否");
            Line(tag + " 含 pic:", xml.IndexOf("pic:", StringComparison.Ordinal) >= 0 ? "是" : "否");
            Line(tag + " 含 w:drawing", xml.IndexOf("w:drawing", StringComparison.Ordinal) >= 0 ? "是" : "否");
            int i = xml.IndexOf("drawing", StringComparison.OrdinalIgnoreCase);
            Line(tag + " 片段", i < 0 ? Snip(xml, 180) : Snip(xml.Substring(Math.Max(0, i - 40)), 220));
        }

        private static int ProbePasteFlavors(WdSession s, string seed)
        {
            string srcKeep = Path.Combine(
                Path.GetTempPath(),
                "ew-wd-src-" + Guid.NewGuid().ToString("N") + ".docx");
            File.Copy(seed, srcKeep, true);
            ToolResult src = s.Open(srcKeep, false, "wps");
            if (!src.Success)
            {
                Line("开源", src.Error ?? "失败");
                return 1;
            }

            bool copied = s.TryWpsDoc(srcKeep, srcDoc =>
            {
                object shape = FirstInline(srcDoc);
                if (shape == null)
                {
                    Line("源", "无 InlineShape");
                    return false;
                }

                try
                {
                    Invoke(GetProp(shape, "Range"), "Copy");
                    return true;
                }
                catch (Exception ex)
                {
                    Line("Copy", TrimErr(ex.Message));
                    return false;
                }
            });
            if (!copied)
            {
                return 1;
            }

            int[] formats = { 0, 3, 8, 10, 13 };
            foreach (int fmt in formats)
            {
                string destPath = s.NewDocxPath();
                ToolResult dest = s.Open(destPath, true, "wps");
                if (!dest.Success)
                {
                    Line("PasteSpecial " + fmt, dest.Error ?? "开空稿失败");
                    continue;
                }

                s.TryWpsDoc(destPath, destDoc =>
                {
                    TrySilenceAlerts(destDoc);
                    object sel = GetProp(GetProp(destDoc, "Application"), "Selection");
                    try
                    {
                        Invoke(sel, "PasteSpecial", fmt);
                        Line("PasteSpecial " + fmt, "成功 " + DescribeCharts(s, destPath));
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            Invoke(GetProp(destDoc, "Content"), "PasteSpecial", fmt);
                            Line("PasteSpecial(content) " + fmt, "成功 " + DescribeCharts(s, destPath));
                        }
                        catch
                        {
                            Line("PasteSpecial " + fmt, TrimErr(ex.Message));
                            return true;
                        }
                    }

                    DumpOpenXmlHints(s, destPath, "PasteSpecial " + fmt + " XML");
                    return true;
                });
            }

            s.CloseDocByPath(srcKeep);
            return 0;
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
