using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using WordAddIn1.PresentationHost;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>同进程反复走产品 apply：铺底建图 + 换数（灌数+钉皮）。</summary>
    internal static class ProbeRepeatLean
    {
        private static readonly string[] Cats = { "A", "B", "C" };
        private static readonly string[] Vals = { "40", "35", "25" };
        private static readonly string[] Cats2 = { "X", "Y", "Z" };
        private static readonly string[] Vals2 = { "10", "20", "70" };

        public static int Run(int rounds, string seedPath, string recycleMode)
        {
            if (string.IsNullOrEmpty(recycleMode))
            {
                recycleMode = "suite";
            }

            Console.WriteLine("探针 · 产品 apply 连跑 lean 标题+饼 换数 rounds=" + rounds
                + " recycle=" + recycleMode
                + " seed=" + (string.IsNullOrEmpty(seedPath) ? "(new)" : seedPath));
            PowerPoint.Application app = null;
            PowerPoint.Presentation dest = null;
            try
            {
                app = new PowerPoint.Application();
                app.Visible = Office.MsoTriState.msoTrue;
                app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                string destPath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(), "ew-dest-punch", "probe-lean-saved.pptx");
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destPath));
                if (!string.IsNullOrEmpty(seedPath) && System.IO.File.Exists(seedPath))
                {
                    System.IO.File.Copy(seedPath, destPath, true);
                    dest = app.Presentations.Open(
                        destPath,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue);
                }
                else
                {
                    dest = app.Presentations.Add(Office.MsoTriState.msoTrue);
                    dest.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                    dest.SaveAs(destPath);
                }

                Ping("start slides=" + dest.Slides.Count, app, dest);

                var session = new ContentGroupSession
                {
                    Presentation = dest,
                    Slide = dest.Slides[1],
                    SlideId = dest.Slides[1].SlideID.ToString(CultureInfo.InvariantCulture)
                };

                for (int i = 1; i <= rounds; i++)
                {
                    if (!TrySeedKeepPages(session, dest, out string err))
                    {
                        Console.WriteLine("R" + i + " SeedDeck 失败: " + err);
                        Ping("R" + i + " seed-fail", app, dest);
                        return DestDead(app, dest) ? 2 : 1;
                    }

                    string create = ContentHtml.MixParts(
                        ContentHtml.MixTitleMarkup("柱数原"),
                        ContentHtml.ChartFrag("pie2d", 8, 22, 50, 50, Cats, Vals));
                    if (!session.TryApply(create, true, out List<string> ids, out err) || ids == null || ids.Count < 1)
                    {
                        Console.WriteLine("R" + i + " 铺底失败: " + err);
                        DumpWarn(session);
                        Ping("R" + i + " create-fail", app, dest);
                        return DestDead(app, dest) ? 2 : 1;
                    }

                    if (!session.TryReadFullPage(out _, out err))
                    {
                        Console.WriteLine("R" + i + " 铺底后读失败: " + err);
                        Ping("R" + i + " read-after-create", app, dest);
                        return DestDead(app, dest) ? 2 : 1;
                    }

                    Ping("R" + i + " after create", app, dest);
                    if (DestDead(app, dest))
                    {
                        return 2;
                    }

                    string chartId = ids[ids.Count - 1];
                    if (!session.TryApply(
                            ContentHtml.ChartUpdate(chartId, Cats2, Vals2, "规模", "pie2d"),
                            false,
                            out _,
                            out err))
                    {
                        Console.WriteLine("R" + i + " 换数失败: " + err);
                        DumpWarn(session);
                        Ping("R" + i + " lean-fail", app, dest);
                        return DestDead(app, dest) ? 2 : 1;
                    }

                    if (!session.TryReadFullPage(out _, out err))
                    {
                        Console.WriteLine("R" + i + " 换数后读失败: " + err);
                        Ping("R" + i + " read-after-lean", app, dest);
                        return DestDead(app, dest) ? 2 : 1;
                    }

                    try
                    {
                        dest.Save();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("R" + i + " Save 失败: " + ex.Message);
                        Ping("R" + i + " save-fail", app, dest);
                        return DestDead(app, dest) ? 2 : 1;
                    }

                    Ping("R" + i + " after lean 页=" + dest.Slides.Count, app, dest);
                    if (DestDead(app, dest))
                    {
                        return 2;
                    }

                    if (i % 6 == 0 && i < rounds)
                    {
                        if (!TryRecycleLikeSuite(ref app, ref dest, destPath, session, recycleMode))
                        {
                            Console.WriteLine("R" + i + " 回收失败");
                            Ping("R" + i + " recycle-fail", app, dest);
                            return 2;
                        }

                        Ping("R" + i + " after recycle", app, dest);
                    }
                }

                Ping("end alive", app, dest);
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("FATAL " + ex);
                Ping("fatal", app, dest);
                return 1;
            }
            finally
            {
                try
                {
                    if (dest != null)
                    {
                        dest.Saved = Office.MsoTriState.msoTrue;
                        dest.Close();
                    }
                }
                catch
                {
                }

                try
                {
                    if (app != null)
                    {
                        app.Quit();
                    }
                }
                catch
                {
                }
            }
        }

        private static bool TrySeedKeepPages(ContentGroupSession s, PowerPoint.Presentation dest, out string error)
        {
            if (!s.TryAddBlankSlide(out string curId, out error) || !s.TrySelectSlide(curId, out error))
            {
                return false;
            }

            if (!s.TryApply(ContentHtml.TextboxCreate("KEEP-CUR", left: 70, top: 4, width: 26, height: 8), true, out _, out error))
            {
                return false;
            }

            if (!s.TryAddBlankSlide(out string a1, out error)
                || !s.TrySelectSlide(a1, out error)
                || !s.TryApply(ContentHtml.TextboxCreate("KEEP-A1", left: 8, top: 8, width: 26, height: 8), true, out _, out error))
            {
                return false;
            }

            if (!s.TryAddBlankSlide(out string a2, out error)
                || !s.TrySelectSlide(a2, out error)
                || !s.TryApply(ContentHtml.TextboxCreate("KEEP-A2", left: 8, top: 8, width: 26, height: 8), true, out _, out error))
            {
                return false;
            }

            return s.TrySelectSlide(a1, out error);
        }

        private static bool TryRecycleLikeSuite(
            ref PowerPoint.Application app,
            ref PowerPoint.Presentation dest,
            string destPath,
            ContentGroupSession session,
            string mode)
        {
            int oldPid = FirstPowerpntPid();
            Console.WriteLine("  recycle mode=" + mode + " pptPid=" + oldPid + " pptN=" + PowerpntCount());
            try
            {
                dest.Save();
                dest.Saved = Office.MsoTriState.msoTrue;
                dest.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("  recycle Close: " + ex.Message);
            }

            dest = null;
            session.Presentation = null;
            session.Slide = null;

            bool sameApp = string.Equals(mode, "sameapp", StringComparison.OrdinalIgnoreCase);
            if (!sameApp)
            {
                try
                {
                    if (app != null && app.Presentations.Count == 0)
                    {
                        app.Quit();
                    }
                }
                catch
                {
                }

                if (!string.Equals(mode, "nofinal", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        if (app != null)
                        {
                            Marshal.FinalReleaseComObject(app);
                        }
                    }
                    catch
                    {
                    }
                }

                app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                if (string.Equals(mode, "waitdead", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(mode, "kill", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.Equals(mode, "kill", StringComparison.OrdinalIgnoreCase) && PidAlive(oldPid))
                    {
                        try
                        {
                            Process.GetProcessById(oldPid).Kill();
                            Console.WriteLine("  recycle killed pid=" + oldPid);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("  recycle kill: " + ex.Message);
                        }
                    }

                    int waited = WaitPowerpntGone(oldPid, 15000);
                    Console.WriteLine("  recycle wait " + waited + "ms still=" + PidAlive(oldPid)
                        + " pptN=" + PowerpntCount());
                }
                else
                {
                    System.Threading.Thread.Sleep(900);
                    Console.WriteLine("  recycle slept 900ms oldPid=" + oldPid
                        + " still=" + PidAlive(oldPid) + " pptN=" + PowerpntCount());
                }
            }

            try
            {
                if (app == null)
                {
                    app = new PowerPoint.Application();
                    app.Visible = Office.MsoTriState.msoTrue;
                    app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                }

                int newPid = FirstPowerpntPid();
                Console.WriteLine("  recycle newApp pptPid=" + newPid
                    + " samePid=" + (newPid != 0 && newPid == oldPid));
                dest = app.Presentations.Open(
                    destPath,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue);
                session.Presentation = dest;
                session.Slide = dest.Slides[dest.Slides.Count];
                session.SlideId = dest.Slides[dest.Slides.Count].SlideID.ToString(CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("  recycle Open: " + ex.Message);
                return false;
            }
        }

        private static int FirstPowerpntPid()
        {
            try
            {
                Process[] ps = Process.GetProcessesByName("POWERPNT");
                return ps.Length == 0 ? 0 : ps[0].Id;
            }
            catch
            {
                return 0;
            }
        }

        private static int PowerpntCount()
        {
            try
            {
                return Process.GetProcessesByName("POWERPNT").Length;
            }
            catch
            {
                return -1;
            }
        }

        private static bool PidAlive(int pid)
        {
            if (pid <= 0)
            {
                return false;
            }

            try
            {
                Process.GetProcessById(pid);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static int WaitPowerpntGone(int oldPid, int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (oldPid <= 0 || !PidAlive(oldPid))
                {
                    return (int)sw.ElapsedMilliseconds;
                }

                System.Threading.Thread.Sleep(200);
            }

            return (int)sw.ElapsedMilliseconds;
        }

        private static bool DestDead(PowerPoint.Application app, PowerPoint.Presentation dest)
        {
            try
            {
                if (dest == null || dest.Application == null || dest.Slides.Count < 1)
                {
                    Console.WriteLine("*** DEST PUNCH app.Presentations=" + TryNames(app));
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("*** DEST PUNCH " + ex.Message + " app.Presentations=" + TryNames(app));
                return true;
            }

            return false;
        }

        private static void Ping(string tag, PowerPoint.Application app, PowerPoint.Presentation dest)
        {
            PptHtmlOoxmlIo.PingDest(null, tag, dest);
            Console.WriteLine("  集合 " + TryNames(app));
        }

        private static string TryNames(PowerPoint.Application app)
        {
            try
            {
                int n = app.Presentations.Count;
                var names = new List<string>();
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        names.Add(app.Presentations[i].Name);
                    }
                    catch
                    {
                        names.Add("[" + i + "死]");
                    }
                }

                return "稿=" + n + " [" + string.Join(",", names.ToArray()) + "]";
            }
            catch (Exception ex)
            {
                return "会话死:" + ex.Message;
            }
        }

        private static void DumpWarn(ContentGroupSession s)
        {
            if (s.LastWarnings == null)
            {
                return;
            }

            foreach (string w in s.LastWarnings)
            {
                if (w != null && (w.IndexOf("打点", StringComparison.Ordinal) >= 0
                    || w.IndexOf("瘦包", StringComparison.Ordinal) >= 0
                    || w.IndexOf("死", StringComparison.Ordinal) >= 0
                    || w.IndexOf("失败", StringComparison.Ordinal) >= 0))
                {
                    Console.WriteLine("  w: " + w);
                }
            }
        }
    }
}
