using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

/// <summary>
/// 拆同进程打穿原稿。产品不 Release 稿对象。死后用一开始握着的 Application 查 Presentations。
/// </summary>
internal static class ProbeDestPunch
{
    public static int Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        string mode = Arg(args, "--mode", "paste");
        int rounds = int.Parse(Arg(args, "--rounds", "30"));
        string seed = Arg(args, "--seed", "");
        if (string.IsNullOrEmpty(seed))
        {
            seed = Path.Combine(AppContext.BaseDirectory, "content-batches", "content-batch-1.pptx");
        }

        string dir = Path.Combine(Path.GetTempPath(), "ew-dest-punch");
        Directory.CreateDirectory(dir);
        string destPath = Path.Combine(dir, "dest-" + mode + ".pptx");
        if (File.Exists(seed))
        {
            File.Copy(seed, destPath, true);
        }

        PowerPoint.Application app = null;
        PowerPoint.Presentation dest = null;
        try
        {
            app = new PowerPoint.Application();
            app.Visible = Office.MsoTriState.msoTrue;
            app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
            dest = File.Exists(destPath)
                ? app.Presentations.Open(destPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue)
                : app.Presentations.Add(Office.MsoTriState.msoTrue);
            if (dest.Slides.Count < 3)
            {
                dest.Slides.Add(dest.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            }

            if (!File.Exists(destPath))
            {
                dest.SaveAs(destPath);
            }

            Dump("start mode=" + mode + " seed=" + (File.Exists(seed) ? "yes" : "no"), app, dest);
            for (int i = 1; i <= rounds; i++)
            {
                if (!RunRound(mode, i, dir, app, dest))
                {
                    return 2;
                }
            }

            Dump("end alive", app, dest);
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine("FATAL " + ex);
            Dump("fatal", app, dest);
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

    private static bool RunRound(
        string mode,
        int i,
        string dir,
        PowerPoint.Application app,
        PowerPoint.Presentation dest)
    {
        string copyPath = Path.Combine(dir, mode + "-" + i + ".pptx");
        dest.SaveCopyAs(copyPath);
        if (Dead("R" + i + " after SaveCopyAs", app, dest))
        {
            return false;
        }

        PowerPoint.Shapes destShapes = dest.Slides[dest.Slides.Count].Shapes;
        bool addChart = mode != "paste-existing";
        bool secondOpen = mode != "paste-once";
        bool doGc = mode == "paste-gc" || mode == "paste-snap";
        bool doSnap = mode == "paste-snap";

        PowerPoint.Presentation copy = OpenCopy(app, copyPath, true);
        if (Dead("R" + i + " after Open1", app, dest))
        {
            return false;
        }

        PowerPoint.Shape src = null;
        if (addChart)
        {
            src = copy.Slides[1].Shapes.AddChart2(
                201, Office.XlChartType.xlPie, 80, 80, 320, 220, true);
            if (Dead("R" + i + " after AddChart", app, dest))
            {
                return false;
            }

            if (doGc)
            {
                TryCloseChartWorkbook(src);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                System.Threading.Thread.Sleep(200);
                if (Dead("R" + i + " after ExcelGC", app, dest))
                {
                    return false;
                }
            }
        }
        else
        {
            src = FindFirstChart(copy) ?? copy.Slides[1].Shapes[1];
        }

        if (!secondOpen)
        {
            src.Copy();
            if (Dead("R" + i + " after Copy1", app, dest))
            {
                return false;
            }
        }

        copy.Saved = Office.MsoTriState.msoTrue;
        copy.Close();
        copy = null;
        if (Dead("R" + i + " after Close1", app, dest))
        {
            return false;
        }

        if (secondOpen)
        {
            copy = OpenCopy(app, copyPath, true);
            if (Dead("R" + i + " after Open2", app, dest))
            {
                return false;
            }

            src = FindFirstChart(copy) ?? copy.Slides[1].Shapes[copy.Slides[1].Shapes.Count];
            src.Copy();
            if (Dead("R" + i + " after Copy2", app, dest))
            {
                return false;
            }

            copy.Saved = Office.MsoTriState.msoTrue;
            copy.Close();
            copy = null;
            if (Dead("R" + i + " after Close2", app, dest))
            {
                return false;
            }
        }

        try
        {
            destShapes.Paste();
        }
        catch (Exception ex)
        {
            Console.WriteLine("  Paste 失败: " + ex.Message);
            try
            {
                destShapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteDefault);
            }
            catch (Exception ex2)
            {
                Console.WriteLine("  PasteSpecial 失败: " + ex2.Message);
            }
        }

        if (Dead("R" + i + " after Paste", app, dest))
        {
            return false;
        }

        if (doSnap)
        {
            TrySnapLastChart(destShapes);
            if (Dead("R" + i + " after Snap", app, dest))
            {
                return false;
            }
        }

        try
        {
            int n = destShapes.Count;
            Dump("R" + i + " destShapes=" + n, app, dest);
        }
        catch (Exception ex)
        {
            Console.WriteLine("*** destShapes 死: " + ex.Message);
            Dump("R" + i + " destShapes-dead", app, dest);
            return false;
        }

        TryDelete(copyPath);
        return true;
    }

    private static PowerPoint.Presentation OpenCopy(PowerPoint.Application app, string path, bool withWindow)
    {
        return app.Presentations.Open(
            path,
            Office.MsoTriState.msoTrue,
            Office.MsoTriState.msoFalse,
            withWindow ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse);
    }

    private static PowerPoint.Shape FindFirstChart(PowerPoint.Presentation pres)
    {
        for (int s = 1; s <= pres.Slides.Count; s++)
        {
            PowerPoint.Shapes shapes = pres.Slides[s].Shapes;
            for (int i = 1; i <= shapes.Count; i++)
            {
                try
                {
                    if (shapes[i].HasChart == Office.MsoTriState.msoTrue)
                    {
                        return shapes[i];
                    }
                }
                catch
                {
                }
            }
        }

        return null;
    }

    private static void TryCloseChartWorkbook(PowerPoint.Shape shape)
    {
        try
        {
            if (shape.HasChart != Office.MsoTriState.msoTrue)
            {
                return;
            }

            object wb = shape.Chart.ChartData.Workbook;
            try
            {
                wb.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, wb, new object[] { false });
            }
            catch
            {
            }

            Marshal.FinalReleaseComObject(wb);
        }
        catch
        {
        }
    }

    private static void TrySnapLastChart(PowerPoint.Shapes shapes)
    {
        try
        {
            for (int i = shapes.Count; i >= 1; i--)
            {
                PowerPoint.Shape s = shapes[i];
                if (s.HasChart != Office.MsoTriState.msoTrue)
                {
                    continue;
                }

                PowerPoint.Chart c = s.Chart;
                object sc = c.SeriesCollection();
                int series = Convert.ToInt32(
                    sc.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, sc, null));
                object fill = c.PlotArea.Format.Fill.Visible;
                Console.WriteLine("  snap series=" + series + " plotFill=" + fill + " type=" + c.ChartType);
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("  snap 失败: " + ex.Message);
        }
    }

    private static bool Dead(string tag, PowerPoint.Application app, PowerPoint.Presentation dest)
    {
        Dump(tag, app, dest);
        string name = TryName(dest);
        object destApp = TryDestApp(dest);
        bool slidesOk = TrySlides(dest) >= 0;
        if (name == null || destApp == null || !slidesOk)
        {
            Console.WriteLine("*** DEST PUNCH name=" + (name ?? "null")
                + " destApp=" + (destApp == null ? "null" : "ok")
                + " slides=" + TrySlides(dest)
                + " app.Presentations=" + TryPresNames(app));
            return true;
        }

        return false;
    }

    private static void Dump(string tag, PowerPoint.Application app, PowerPoint.Presentation dest)
    {
        int excel = 0;
        try
        {
            excel = Process.GetProcessesByName("EXCEL").Length;
        }
        catch
        {
        }

        string destBit = "dest=null";
        if (dest != null)
        {
            destBit = "name=" + (TryName(dest) ?? "死")
                + " slides=" + TrySlides(dest)
                + " destApp=" + (TryDestApp(dest) == null ? "null" : "ok");
        }

        Console.WriteLine(tag + " | " + destBit + " | " + TryPresNames(app) + " | excel=" + excel);
    }

    private static string TryPresNames(PowerPoint.Application app)
    {
        if (app == null)
        {
            return "app=null";
        }

        try
        {
            int n = app.Presentations.Count;
            var names = new System.Collections.Generic.List<string>();
            for (int i = 1; i <= n; i++)
            {
                try
                {
                    names.Add(app.Presentations[i].Name);
                }
                catch (Exception ex)
                {
                    names.Add("[" + i + "死:" + ex.GetType().Name + "]");
                }
            }

            return "稿=" + n + " [" + string.Join(",", names.ToArray()) + "]";
        }
        catch (Exception ex)
        {
            return "会话死:" + ex.Message;
        }
    }

    private static string TryName(PowerPoint.Presentation dest)
    {
        try
        {
            return dest == null ? null : dest.Name;
        }
        catch
        {
            return null;
        }
    }

    private static object TryDestApp(PowerPoint.Presentation dest)
    {
        try
        {
            return dest == null ? null : dest.Application;
        }
        catch
        {
            return null;
        }
    }

    private static int TrySlides(PowerPoint.Presentation dest)
    {
        try
        {
            return dest == null ? -1 : dest.Slides.Count;
        }
        catch
        {
            return -2;
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
        }
    }

    private static string Arg(string[] args, string key, string fallback)
    {
        foreach (string a in args)
        {
            if (a.StartsWith(key + "=", StringComparison.OrdinalIgnoreCase))
            {
                return a.Substring(key.Length + 1);
            }
        }

        return fallback;
    }
}
