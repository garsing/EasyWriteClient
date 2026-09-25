using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

internal static class ProbeSessionDirt
{
    private static int _excelPeak;

    public static int Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        string dir = Path.Combine(Path.GetTempPath(), "ew-session-dirt");
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
            dest.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            dest.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
            dest.Slides.Add(3, PowerPoint.PpSlideLayout.ppLayoutBlank);
            dest.SaveAs(destPath);
            Dump("start", app, dest);

            int rounds = 20;
            for (int i = 1; i <= rounds; i++)
            {
                string copyPath = Path.Combine(dir, "copy-" + i + ".pptx");
                dest.SaveCopyAs(copyPath);
                Dump("R" + i + " after SaveCopyAs", app, dest);

                PowerPoint.Presentation copy = app.Presentations.Open(
                    copyPath,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue);
                Dump("R" + i + " after Open bypass", app, dest);

                PowerPoint.Shape chartShape = copy.Slides[2].Shapes.AddChart2(
                    201, Microsoft.Office.Core.XlChartType.xlColumnClustered,
                    80, 80, 400, 250, true);
                Dump("R" + i + " after AddChart2", app, dest);

                try
                {
                    if (chartShape.HasChart == Office.MsoTriState.msoTrue)
                    {
                        object wb = chartShape.Chart.ChartData.Workbook;
                        object excelApp = wb.GetType().InvokeMember(
                            "Application",
                            System.Reflection.BindingFlags.GetProperty,
                            null,
                            wb,
                            null);
                        excelApp.GetType().InvokeMember(
                            "Quit",
                            System.Reflection.BindingFlags.InvokeMethod,
                            null,
                            excelApp,
                            null);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  quit-excel: " + ex.Message);
                }

                copy.Save();
                copy.Saved = Office.MsoTriState.msoTrue;
                copy.Close();
                Marshal.ReleaseComObject(copy);
                copy = null;
                Dump("R" + i + " after Close bypass", app, dest);

                if (!DestInCollection(app, dest))
                {
                    Console.WriteLine("*** DEST LEFT Presentations after Close bypass R" + i);
                    break;
                }

                PowerPoint.Presentation slim = app.Presentations.Open(
                    copyPath,
                    Office.MsoTriState.msoTrue,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue);
                Dump("R" + i + " after Open 2nd (slim-like)", app, dest);
                slim.Saved = Office.MsoTriState.msoTrue;
                slim.Close();
                Marshal.ReleaseComObject(slim);
                slim = null;
                Dump("R" + i + " after Close 2nd", app, dest);

                if (!DestInCollection(app, dest) || !DestAlive(dest))
                {
                    Console.WriteLine("*** DEST DEAD after Open/Close 2nd R" + i);
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
            try
            {
                if (dest != null)
                {
                    dest.Saved = Office.MsoTriState.msoTrue;
                    dest.Close();
                }
            }
            catch { }
            try { if (app != null) app.Quit(); } catch { }
        }
    }

    private static bool DestAlive(PowerPoint.Presentation dest)
    {
        try
        {
            return dest != null && !string.IsNullOrEmpty(dest.Name);
        }
        catch
        {
            return false;
        }
    }

    private static bool DestInCollection(PowerPoint.Application app, PowerPoint.Presentation dest)
    {
        string destName;
        string destFull;
        try
        {
            destName = dest.Name;
            destFull = dest.FullName;
        }
        catch
        {
            return false;
        }

        try
        {
            for (int i = 1; i <= app.Presentations.Count; i++)
            {
                try
                {
                    PowerPoint.Presentation one = app.Presentations[i];
                    if (string.Equals(one.FullName, destFull, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(one.Name, destName, StringComparison.OrdinalIgnoreCase))
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

    private static void Dump(string tag, PowerPoint.Application app, PowerPoint.Presentation dest)
    {
        int excel = 0;
        try { excel = Process.GetProcessesByName("EXCEL").Length; } catch { }
        if (excel > _excelPeak) _excelPeak = excel;
        int ppt = 0;
        try { ppt = Process.GetProcessesByName("POWERPNT").Length; } catch { }

        string destBit = "dest=null";
        if (dest != null)
        {
            try
            {
                destBit = "dest=活(" + dest.Name + ") win=" + dest.Windows.Count
                    + " inCol=" + DestInCollection(app, dest);
            }
            catch (Exception ex)
            {
                destBit = "dest=死:" + ex.Message;
            }
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

                int win = -1;
                try { win = app.Windows.Count; } catch { }
                sess = "稿=" + n + " [" + string.Join(",", names) + "] appWin=" + win;
            }
            catch (Exception ex)
            {
                sess = "会话读失败:" + ex.Message;
            }
        }

        Console.WriteLine(tag + " | " + destBit + " | " + sess
            + " | excel=" + excel + " peak=" + _excelPeak + " ppt=" + ppt);
    }
}
