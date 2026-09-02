using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    internal static class PptHtmlGroupIo
    {
        public static bool TryGroupPowerPoint(
            PowerPoint.Slide slide,
            IList<PowerPoint.Shape> members,
            out PowerPoint.Shape group,
            out string error)
        {
            group = null;
            error = null;
            if (slide == null || members == null || members.Count == 0)
            {
                error = "Group 需要至少一个子形状";
                return false;
            }

            try
            {
                if (members.Count == 1)
                {
                    PowerPoint.Shape only = members[0];
                    PowerPoint.Shape backing = slide.Shapes.AddShape(
                        Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                        only.Left,
                        only.Top,
                        Math.Max(1f, only.Width),
                        Math.Max(1f, only.Height));
                    try
                    {
                        backing.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                        backing.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                    }
                    catch (Exception)
                    {
                    }

                    members = new List<PowerPoint.Shape> { backing, only };
                }

                var names = new object[members.Count];
                for (int i = 0; i < members.Count; i++)
                {
                    names[i] = members[i].Name;
                }

                group = slide.Shapes.Range(names).Group();
                return group != null;
            }
            catch (Exception ex)
            {
                error = "COM Group 失败: " + ex.Message;
                return false;
            }
        }

        public static bool TryGroupWpp(
            object slide,
            IList<object> members,
            out object group,
            out string error)
        {
            group = null;
            error = null;
            if (slide == null || members == null || members.Count == 0)
            {
                error = "Group 需要至少一个子形状";
                return false;
            }

            try
            {
                object shapes = WppCom.GetProperty(slide, "Shapes");
                if (members.Count == 1)
                {
                    object only = members[0];
                    double left = Convert.ToDouble(WppCom.GetProperty(only, "Left"));
                    double top = Convert.ToDouble(WppCom.GetProperty(only, "Top"));
                    double width = Math.Max(1.0, Convert.ToDouble(WppCom.GetProperty(only, "Width")));
                    double height = Math.Max(1.0, Convert.ToDouble(WppCom.GetProperty(only, "Height")));
                    object backing = WppCom.Invoke(
                        shapes,
                        "AddShape",
                        1,
                        left,
                        top,
                        width,
                        height);
                    if (backing != null)
                    {
                        try
                        {
                            object fill = WppCom.GetProperty(backing, "Fill");
                            WppCom.TrySetProperty(fill, "Visible", 0);
                            object line = WppCom.GetProperty(backing, "Line");
                            WppCom.TrySetProperty(line, "Visible", 0);
                        }
                        catch (Exception)
                        {
                        }

                        members = new List<object> { backing, only };
                    }
                }

                var names = new object[members.Count];
                for (int i = 0; i < members.Count; i++)
                {
                    names[i] = WppCom.GetProperty(members[i], "Name");
                }

                object range = WppCom.Invoke(shapes, "Range", new object[] { names });
                group = WppCom.Invoke(range, "Group");
                return group != null;
            }
            catch (Exception ex)
            {
                error = "WPP Group 失败: " + ex.Message;
                return false;
            }
        }
    }
}
