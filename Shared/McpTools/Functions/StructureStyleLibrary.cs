using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// styles.xml 里还没有、样式库里已经有的内置样式，写进当前稿的样式定义。不套到正文上。
    /// </summary>
    internal static class StructureStyleLibrary
    {
        public static XElement Ensure(Word.Document doc, XElement stylesRoot)
        {
            if (doc == null)
            {
                return stylesRoot;
            }

            var present = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Collect(stylesRoot, present);
            bool changed = false;
            Word.Application app = null;
            bool screen = true;
            Word.WdAlertLevel alerts = Word.WdAlertLevel.wdAlertsNone;
            try
            {
                app = doc.Application;
                screen = app.ScreenUpdating;
                alerts = app.DisplayAlerts;
                app.ScreenUpdating = false;
                app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            }
            catch (Exception)
            {
                app = null;
            }

            try
            {
                foreach (Word.Style style in doc.Styles)
                {
                    try
                    {
                        if (!style.BuiltIn)
                        {
                            continue;
                        }

                        Word.WdStyleType type = style.Type;
                        if (type != Word.WdStyleType.wdStyleTypeParagraph
                            && type != Word.WdStyleType.wdStyleTypeCharacter)
                        {
                            continue;
                        }

                        string local = style.NameLocal;
                        if (string.IsNullOrEmpty(local) || Known(present, local))
                        {
                            continue;
                        }

                        if (!Touch(style))
                        {
                            continue;
                        }

                        changed = true;
                        Remember(present, local);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            finally
            {
                if (app != null)
                {
                    try
                    {
                        app.DisplayAlerts = alerts;
                        app.ScreenUpdating = screen;
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            if (!changed)
            {
                return stylesRoot;
            }

            OpenXmlPackage full = OpenXmlPackage.LoadFull(doc);
            XElement next = full.FindPartRoot("styles.xml", OpenXmlPackage.W + "styles");
            return next ?? stylesRoot;
        }

        private static void Collect(XElement stylesRoot, HashSet<string> present)
        {
            if (stylesRoot == null)
            {
                return;
            }

            foreach (XElement style in stylesRoot.Elements(OpenXmlPackage.W + "style"))
            {
                string name = (string)style.Element(OpenXmlPackage.W + "name")?.Attribute(OpenXmlPackage.W + "val");
                Remember(present, name);
            }
        }

        private static void Remember(HashSet<string> present, string name)
        {
            foreach (string item in StructureStyleCatalog.EquivalentNames(name))
            {
                present.Add(item);
            }
        }

        private static bool Known(HashSet<string> present, string name)
        {
            foreach (string item in StructureStyleCatalog.EquivalentNames(name))
            {
                if (present.Contains(item))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>把基准名再写一次。必须传样式名，传样式对象时 WPS 会多出一支叫 macro 的样式。</summary>
        private static bool Touch(Word.Style style)
        {
            try
            {
                object based = style.get_BaseStyle();
                string name = null;
                Word.Style basedStyle = based as Word.Style;
                if (basedStyle != null)
                {
                    name = basedStyle.NameLocal;
                }
                else if (based != null)
                {
                    name = Convert.ToString(based);
                }

                if (!string.IsNullOrEmpty(name))
                {
                    object arg = name;
                    style.set_BaseStyle(ref arg);
                    return true;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                int priority = style.Priority;
                style.Priority = priority;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
