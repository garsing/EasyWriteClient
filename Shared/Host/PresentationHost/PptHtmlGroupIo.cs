using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    internal sealed class PptHtmlGroupXmlChild
    {
        public int Id;
        public bool IsGroup;

        public static bool HasGroup(IList<PptHtmlGroupXmlChild> kids)
        {
            if (kids == null)
            {
                return false;
            }

            for (int i = 0; i < kids.Count; i++)
            {
                if (kids[i].IsGroup)
                {
                    return true;
                }
            }

            return false;
        }
    }

    internal static partial class PptHtmlGroupIo
    {
        private const int MsoGroup = 6;

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
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        only.Left,
                        only.Top,
                        Math.Max(1f, only.Width),
                        Math.Max(1f, only.Height));
                    try
                    {
                        backing.Fill.Visible = Office.MsoTriState.msoFalse;
                        backing.Line.Visible = Office.MsoTriState.msoFalse;
                    }
                    catch (Exception)
                    {
                    }

                    members = new List<PowerPoint.Shape> { backing, only };
                }

                return TryGroupPowerPointMembers(slide, members, out group, out error);
            }
            catch (Exception ex)
            {
                error = "COM Group 失败: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 管形状 / 新建组共用。成员都是叶子走 Range.Group；已有组则 OOXML 包一层，避免拍平。
        /// </summary>
        public static bool TryGroupPowerPointMembers(
            PowerPoint.Slide slide,
            IList<PowerPoint.Shape> members,
            out PowerPoint.Shape group,
            out string error)
        {
            group = null;
            error = null;
            if (slide == null || members == null || members.Count < 2)
            {
                error = "Group 需要至少两个子形状";
                return false;
            }

            var names = new object[members.Count];
            bool hasInnerGroup = false;
            for (int i = 0; i < members.Count; i++)
            {
                if (members[i] == null)
                {
                    error = "Group 成员不能为空";
                    return false;
                }

                names[i] = members[i].Name;
                if (IsGroupShape(members[i]))
                {
                    hasInnerGroup = true;
                }
            }

            try
            {
                if (!hasInnerGroup)
                {
                    group = slide.Shapes.Range(names).Group();
                    return group != null;
                }

                return TryGroupNestedByOoxmlPpt(slide, members, out group, out error);
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

                return TryGroupWppMembers(slide, members, out group, out error);
            }
            catch (Exception ex)
            {
                error = "WPP Group 失败: " + ex.Message;
                return false;
            }
        }

        public static bool TryGroupWppMembers(
            object slide,
            IList<object> members,
            out object group,
            out string error)
        {
            group = null;
            error = null;
            if (slide == null || members == null || members.Count < 2)
            {
                error = "Group 需要至少两个子形状";
                return false;
            }

            object shapes = WppCom.GetProperty(slide, "Shapes");
            if (shapes == null)
            {
                error = "无法读取 Shapes";
                return false;
            }

            var names = new object[members.Count];
            bool hasInnerGroup = false;
            for (int i = 0; i < members.Count; i++)
            {
                if (members[i] == null)
                {
                    error = "Group 成员不能为空";
                    return false;
                }

                names[i] = WppCom.GetProperty(members[i], "Name");
                if (IsWppGroup(members[i]))
                {
                    hasInnerGroup = true;
                }
            }

            try
            {
                if (!hasInnerGroup)
                {
                    object range = WppCom.Invoke(shapes, "Range", new object[] { names });
                    if (range == null)
                    {
                        error = "WPP Range 失败";
                        return false;
                    }

                    group = WppCom.Invoke(range, "Group");
                    return group != null;
                }

                return TryGroupNestedByOoxmlWpp(slide, members, out group, out error);
            }
            catch (Exception ex)
            {
                error = "WPP Group 失败: " + ex.Message;
                return false;
            }
        }

        public static bool TryCopyGroupPowerPoint(PowerPoint.Slide slide, int groupId, out string error)
        {
            error = null;
            if (slide == null || groupId < 1)
            {
                error = "拷组：slide/id 无效";
                return false;
            }

            PowerPoint.Shape top = FindTopLevelPpt(slide.Shapes, groupId);
            if (IsGroupShape(top))
            {
                top.Copy();
                return true;
            }

            return TryCopyInnerGroupByUngroupPpt(slide, groupId, out error);
        }

        public static bool TryCopyGroupWpp(object slide, int groupId, out string error)
        {
            error = null;
            if (slide == null || groupId < 1)
            {
                error = "拷组：slide/id 无效";
                return false;
            }

            object shapes = WppCom.GetProperty(slide, "Shapes");
            object top = FindTopLevelWpp(shapes, groupId);
            if (IsWppGroup(top))
            {
                WppCom.Invoke(top, "Copy");
                return true;
            }

            return TryCopyInnerGroupByUngroupWpp(slide, groupId, out error);
        }

        private static bool IsGroupShape(PowerPoint.Shape shape)
        {
            try
            {
                return shape != null && (int)shape.Type == MsoGroup;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool IsWppGroup(object shape)
        {
            try
            {
                object type = WppCom.GetProperty(shape, "Type");
                return type != null && Convert.ToInt32(type) == MsoGroup;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
