using System;
using System.Collections.Generic;
using System.Globalization;
using WordAddIn1.OpenFiles;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace WordAddIn1.PresentationHost
{
    internal sealed class PptPaletteCollectResult
    {
        public List<string> Palette { get; set; }

        public bool Sampled { get; set; }

        public int ScannedCount { get; set; }
    }

    /// <summary>本演示文稿允许色表。只允许概览 Collect；其它工具 Peek 缓存，禁止再扫页。</summary>
    internal static class PptPaletteIo
    {
        public const int UsedSlideCap = 50;

        private static readonly object CacheSync = new object();
        private static readonly Dictionary<string, List<string>> Cache =
            new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        public static PptPaletteCollectResult CollectPowerPoint(PowerPoint.Presentation presentation)
        {
            var colors = new HashSet<string>(StringComparer.Ordinal);
            AddAlwaysLegal(colors);
            AddThemePowerPoint(presentation, colors);

            int slideCount = 0;
            try
            {
                slideCount = presentation.Slides.Count;
            }
            catch (Exception)
            {
            }

            List<int> indexes = PickSlideIndexes(slideCount, UsedSlideCap);
            foreach (int index in indexes)
            {
                try
                {
                    CollectUsedPowerPoint(presentation.Slides[index], colors);
                }
                catch (Exception)
                {
                }
            }

            PptPaletteCollectResult result = Finish(colors, slideCount, indexes.Count);
            Remember(KeyPowerPoint(presentation), result.Palette);
            return result;
        }

        public static PptPaletteCollectResult CollectWpp(object presentation)
        {
            var colors = new HashSet<string>(StringComparer.Ordinal);
            AddAlwaysLegal(colors);
            AddThemeWpp(presentation, colors);

            object slides = null;
            int slideCount = 0;
            try
            {
                slides = WppCom.GetProperty(presentation, "Slides");
                slideCount = Convert.ToInt32(WppCom.GetProperty(slides, "Count"));
            }
            catch (Exception)
            {
            }

            List<int> indexes = PickSlideIndexes(slideCount, UsedSlideCap);
            if (slides != null)
            {
                foreach (int index in indexes)
                {
                    try
                    {
                        CollectUsedWpp(WppCom.GetIndexed(slides, index), colors);
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            PptPaletteCollectResult result = Finish(colors, slideCount, indexes.Count);
            Remember(KeyWpp(presentation), result.Palette);
            return result;
        }

        /// <summary>上次概览扫过的色表；从未概览则 null。不碰 COM。</summary>
        public static List<string> PeekCachedPowerPoint(PowerPoint.Presentation presentation)
        {
            return Peek(KeyPowerPoint(presentation));
        }

        /// <summary>上次概览扫过的色表；从未概览则 null。不碰 COM。</summary>
        public static List<string> PeekCachedWpp(object presentation)
        {
            return Peek(KeyWpp(presentation));
        }

        public static bool IsAlwaysLegal(string normalized)
        {
            if (string.IsNullOrEmpty(normalized))
            {
                return false;
            }

            return string.Equals(normalized, "none", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "#000000", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "#FFFFFF", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsInPalette(ICollection<string> palette, string normalized)
        {
            if (IsAlwaysLegal(normalized))
            {
                return true;
            }

            if (palette == null || string.IsNullOrEmpty(normalized))
            {
                return false;
            }

            foreach (string one in palette)
            {
                if (string.Equals(one, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        internal static List<int> PickSlideIndexes(int slideCount, int cap)
        {
            var list = new List<int>();
            if (slideCount <= 0)
            {
                return list;
            }

            if (slideCount <= cap)
            {
                for (int i = 1; i <= slideCount; i++)
                {
                    list.Add(i);
                }

                return list;
            }

            var set = new SortedSet<int> { 1, slideCount };
            int slots = cap - 2;
            if (slots > 0)
            {
                for (int i = 1; i <= slots; i++)
                {
                    int idx = 1 + (int)Math.Round(i * (slideCount - 1) / (double)(slots + 1));
                    if (idx < 1)
                    {
                        idx = 1;
                    }

                    if (idx > slideCount)
                    {
                        idx = slideCount;
                    }

                    set.Add(idx);
                }
            }

            int cursor = 2;
            while (set.Count < cap && cursor < slideCount)
            {
                set.Add(cursor);
                cursor++;
            }

            list.AddRange(set);
            if (list.Count > cap)
            {
                list = list.GetRange(0, cap);
                if (!list.Contains(slideCount))
                {
                    list[list.Count - 1] = slideCount;
                    list.Sort();
                }
            }

            return list;
        }

        private static PptPaletteCollectResult Finish(HashSet<string> colors, int slideCount, int scanned)
        {
            var list = new List<string>(colors);
            list.Sort(StringComparer.Ordinal);
            return new PptPaletteCollectResult
            {
                Palette = list,
                Sampled = slideCount > UsedSlideCap,
                ScannedCount = scanned
            };
        }

        private static void AddAlwaysLegal(HashSet<string> colors)
        {
            colors.Add("#000000");
            colors.Add("#FFFFFF");
            colors.Add("none");
        }

        private static void AddThemePowerPoint(PowerPoint.Presentation presentation, HashSet<string> colors)
        {
            try
            {
                Office.ThemeColorScheme scheme = presentation.SlideMaster.Theme.ThemeColorScheme;
                for (int i = 1; i <= 12; i++)
                {
                    try
                    {
                        int rgb = scheme.Colors((Office.MsoThemeColorSchemeIndex)i).RGB;
                        colors.Add(PptHtmlStyleIo.FormatOfficeRgb(rgb));
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void AddThemeWpp(object presentation, HashSet<string> colors)
        {
            try
            {
                object master = WppCom.GetProperty(presentation, "SlideMaster");
                object theme = master == null ? null : WppCom.GetProperty(master, "Theme");
                object scheme = theme == null ? null : WppCom.GetProperty(theme, "ThemeColorScheme");
                if (scheme == null)
                {
                    return;
                }

                object colorsCol = WppCom.GetProperty(scheme, "Colors");
                for (int i = 1; i <= 12; i++)
                {
                    try
                    {
                        object slot = colorsCol == null
                            ? WppCom.Invoke(scheme, "Colors", i)
                            : WppCom.GetIndexed(colorsCol, i);
                        object rgb = slot == null ? null : WppCom.GetProperty(slot, "RGB");
                        if (rgb != null)
                        {
                            colors.Add(PptHtmlStyleIo.FormatOfficeRgb(Convert.ToInt32(rgb)));
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void CollectUsedPowerPoint(PowerPoint.Slide slide, HashSet<string> colors)
        {
            if (slide == null)
            {
                return;
            }

            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                try
                {
                    AddSolidFill(shape.Fill, colors);
                }
                catch (Exception)
                {
                }

                try
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        int rgb = shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB;
                        colors.Add(PptHtmlStyleIo.FormatOfficeRgb(rgb));
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        private static void AddSolidFill(PowerPoint.FillFormat fill, HashSet<string> colors)
        {
            if (fill == null)
            {
                return;
            }

            try
            {
                if (fill.Visible != Office.MsoTriState.msoTrue)
                {
                    colors.Add("none");
                    return;
                }

                colors.Add(PptHtmlStyleIo.FormatOfficeRgb(fill.ForeColor.RGB));
            }
            catch (Exception)
            {
            }
        }

        private static void CollectUsedWpp(object slide, HashSet<string> colors)
        {
            if (slide == null)
            {
                return;
            }

            object shapes = WppCom.GetProperty(slide, "Shapes");
            if (shapes == null)
            {
                return;
            }

            int count;
            try
            {
                count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count"));
            }
            catch (Exception)
            {
                return;
            }

            for (int i = 1; i <= count; i++)
            {
                object shape = WppCom.GetIndexed(shapes, i);
                try
                {
                    object fill = WppCom.GetProperty(shape, "Fill");
                    object visible = fill == null ? null : WppCom.GetProperty(fill, "Visible");
                    object type = fill == null ? null : WppCom.GetProperty(fill, "Type");
                    if (visible != null && Convert.ToInt32(visible) == 0)
                    {
                        colors.Add("none");
                    }
                    else if (type != null && Convert.ToInt32(type) == 1)
                    {
                        object fore = WppCom.GetProperty(fill, "ForeColor");
                        object rgb = fore == null ? null : WppCom.GetProperty(fore, "RGB");
                        if (rgb != null)
                        {
                            colors.Add(PptHtmlStyleIo.FormatOfficeRgb(Convert.ToInt32(rgb)));
                        }
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object tf2 = WppCom.GetProperty(shape, "TextFrame2");
                    object tr = tf2 == null ? null : WppCom.GetProperty(tf2, "TextRange");
                    object font = tr == null ? null : WppCom.GetProperty(tr, "Font");
                    object ffill = font == null ? null : WppCom.GetProperty(font, "Fill");
                    object fore = ffill == null ? null : WppCom.GetProperty(ffill, "ForeColor");
                    object rgb = fore == null ? null : WppCom.GetProperty(fore, "RGB");
                    if (rgb != null)
                    {
                        colors.Add(PptHtmlStyleIo.FormatOfficeRgb(Convert.ToInt32(rgb)));
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        public static string NewColorWarning(string shapeId, string hex)
        {
            string id = string.IsNullOrEmpty(shapeId) ? "new" : shapeId;
            return id + " 用了新的颜色 " + (hex ?? "");
        }

        private static void Remember(string key, List<string> palette)
        {
            if (string.IsNullOrEmpty(key) || palette == null)
            {
                return;
            }

            var copy = new List<string>(palette);
            lock (CacheSync)
            {
                Cache[key] = copy;
            }
        }

        private static List<string> Peek(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                return null;
            }

            lock (CacheSync)
            {
                if (!Cache.TryGetValue(key, out List<string> cached) || cached == null)
                {
                    return null;
                }

                return new List<string>(cached);
            }
        }

        private static string KeyPowerPoint(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return null;
            }

            try
            {
                string full = presentation.FullName;
                if (!string.IsNullOrWhiteSpace(full))
                {
                    return "ppt:" + full;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                string name = presentation.Name;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    return "ppt:" + name;
                }
            }
            catch (Exception)
            {
            }

            return "ppt:unknown";
        }

        private static string KeyWpp(object presentation)
        {
            if (presentation == null)
            {
                return null;
            }

            try
            {
                object full = WppCom.GetProperty(presentation, "FullName");
                if (full != null && !string.IsNullOrWhiteSpace(Convert.ToString(full)))
                {
                    return "wpp:" + Convert.ToString(full);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object name = WppCom.GetProperty(presentation, "Name");
                if (name != null && !string.IsNullOrWhiteSpace(Convert.ToString(name)))
                {
                    return "wpp:" + Convert.ToString(name);
                }
            }
            catch (Exception)
            {
            }

            return "wpp:unknown";
        }
    }
}
