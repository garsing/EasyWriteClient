using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

internal static class ProbeSessionDirtDest
{
    public static int Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        string dir = Path.Combine(Path.GetTempPath(), "ew-session-dirt-dest");
        Directory.CreateDirectory(dir);
        string destPath = Path.Combine(dir, "dest.pptx");
        PowerPoint.Application app = null;
        PowerPoint.Presentation dest = null;
        try
        {
            app = new PowerPoint.Application();
            app.Visible = Office.MsoTriState.msoTrue;
            app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
            dest = app.Presentations.Add(Office.MsoTriState.msoTrue);
            for (int s = 1; s <= 3; s++)
            {
                dest.Slides.Add(s, PowerPoint.PpSlideLayout.ppLayoutBlank);
            }

            dest.SaveAs(destPath);
            Dump("start", app, dest);

            for (int i = 1; i <= 16; i++)
            {
                dest.Slides.Add(dest.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                dest.Slides.Add(dest.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                dest.Slides.Add(dest.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                PowerPoint.Slide live = dest.Slides[dest.Slides.Count - 1];
                PowerPoint.Shape onDest = live.Shapes.AddChart2(
                    201, Microsoft.Office.Core.XlChartType.xlColumnClustered,
                    80, 80, 400, 250, true);
                string snap = WalkChart(onDest);
                Dump("R" + i + " after dest AddChart+snap " + snap, app, dest);

                string copyPath = Path.Combine(dir, "copy-" + i + ".pptx");
                dest.SaveCopyAs(copyPath);
                PowerPoint.Presentation copy = app.Presentations.Open(
                    copyPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
                copy.Slides[copy.Slides.Count - 1].Shapes.AddChart2(
                    201, Microsoft.Office.Core.XlChartType.xlColumnClustered,
                    80, 80, 400, 250, true);
                copy.Save();
                copy.Saved = Office.MsoTriState.msoTrue;
                copy.Close();
                Dump("R" + i + " after Close bypass", app, dest);

                if (!InCol(app, dest) || !Alive(dest))
                {
                    Console.WriteLine("*** DEST GONE after Close bypass R" + i);
                    break;
                }

                PowerPoint.Presentation second = app.Presentations.Open(
                    copyPath, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
                second.Saved = Office.MsoTriState.msoTrue;
                second.Close();
                Dump("R" + i + " after Close 2nd", app, dest);
                if (!InCol(app, dest) || !Alive(dest))
                {
                    Console.WriteLine("*** DEST GONE after Close 2nd R" + i);
                    break;
                }

                try { File.Delete(copyPath); } catch { }
            }

            Dump("end", app, dest);
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
            try { if (dest != null) { dest.Saved = Office.MsoTriState.msoTrue; dest.Close(); } } catch { }
            try { if (app != null) app.Quit(); } catch { }
        }
    }

    private static string WalkChart(PowerPoint.Shape shape)
    {
        try
        {
            if (shape.HasChart != Office.MsoTriState.msoTrue)
            {
                return "noChart";
            }

            PowerPoint.Chart ch = shape.Chart;
            object fill = ch.PlotArea.Format.Fill;
            object type = fill.GetType().InvokeMember("Type", System.Reflection.BindingFlags.GetProperty, null, fill, null);
            object stops = null;
            try
            {
                stops = fill.GetType().InvokeMember("GradientStops", System.Reflection.BindingFlags.GetProperty, null, fill, null);
            }
            catch (Exception ex)
            {
                return "SNAP_FAIL:" + ex.Message;
            }

            return "plotFillType=" + type + " stops=" + (stops == null ? "null" : "ok");
        }
        catch (Exception ex)
        {
            return "SNAP_FAIL:" + ex.Message;
        }
    }

    private static bool Alive(PowerPoint.Presentation dest)
    {
        try { return dest != null && !string.IsNullOrEmpty(dest.Name); }
        catch { return false; }
    }

    private static bool InCol(PowerPoint.Application app, PowerPoint.Presentation dest)
    {
        string full;
        try { full = dest.FullName; }
        catch { return false; }
        try
        {
            for (int i = 1; i <= app.Presentations.Count; i++)
            {
                try
                {
                    if (string.Equals(app.Presentations[i].FullName, full, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                catch { }
            }
        }
        catch { }
        return false;
    }

    private static void Dump(string tag, PowerPoint.Application app, PowerPoint.Presentation dest)
    {
        int excel = 0, ppt = 0;
        try { excel = Process.GetProcessesByName("EXCEL").Length; } catch { }
        try { ppt = Process.GetProcessesByName("POWERPNT").Length; } catch { }
        string destBit = "dest=null";
        if (dest != null)
        {
            try { destBit = "dest=活 win=" + dest.Windows.Count + " inCol=" + InCol(app, dest) + " slides=" + dest.Slides.Count; }
            catch (Exception ex) { destBit = "dest=死:" + ex.Message; }
        }

        string sess = "app=null";
        if (app != null)
        {
            try
            {
                int n = app.Presentations.Count;
                var names = new List<string>();
                for (int i = 1; i <= n; i++)
                {
                    try { names.Add(app.Presentations[i].Name); }
                    catch (Exception ex) { names.Add("[" + i + "死:" + ex.Message + "]"); }
                }

                sess = "稿=" + n + " [" + string.Join(",", names) + "]";
            }
            catch (Exception ex) { sess = "会话读失败:" + ex.Message; }
        }

        Console.WriteLine(tag + " | " + destBit + " | " + sess + " | excel=" + excel + " ppt=" + ppt);
    }
}
