using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1.PresentationHost
{
    /// <summary>页内 table 形状 COM 读写（PPT Interop；WPP 见 .Wpp）。</summary>
    internal static partial class PptHtmlTableIo
    {
        public static string BuildInnerHtml(PptHtmlTableGrid grid)
        {
            var sb = new StringBuilder();
            PptHtmlTableParse.AppendInnerHtml(sb, grid, "    ");
            return sb.ToString();
        }

        private static float[] NormalizePcts(float[] pcts)
        {
            float sum = 0;
            for (int i = 0; i < pcts.Length; i++)
            {
                sum += pcts[i];
            }

            if (sum <= 0)
            {
                float even = 100f / pcts.Length;
                var evenArr = new float[pcts.Length];
                for (int i = 0; i < evenArr.Length; i++)
                {
                    evenArr[i] = even;
                }

                return evenArr;
            }

            if (Math.Abs(sum - 100f) <= 0.5f)
            {
                return pcts;
            }

            var norm = new float[pcts.Length];
            for (int i = 0; i < pcts.Length; i++)
            {
                norm[i] = pcts[i] * 100f / sum;
            }

            return norm;
        }

        private static bool TryParseRgb(string color, out int rgb)
        {
            rgb = 0;
            if (string.IsNullOrWhiteSpace(color) || color.Length != 7 || color[0] != '#')
            {
                return false;
            }

            if (!int.TryParse(color.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int r)
                || !int.TryParse(color.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int g)
                || !int.TryParse(color.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int b))
            {
                return false;
            }

            rgb = r + (g << 8) + (b << 16);
            return true;
        }

        private static string FormatRgb(int rgb)
        {
            int r = rgb & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;
            return "#" + r.ToString("X2") + g.ToString("X2") + b.ToString("X2");
        }

        private static int Pack(int r, int c)
        {
            return (r << 16) | c;
        }

        private static void Unpack(int key, out int r, out int c)
        {
            r = key >> 16;
            c = key & 0xFFFF;
        }

        private sealed class MergeSpan
        {
            public int RowSpan;
            public int ColSpan;
        }

        private sealed class MergeMap
        {
            public Dictionary<int, MergeSpan> Anchors = new Dictionary<int, MergeSpan>();
            public HashSet<int> Covered = new HashSet<int>();
        }
    }
}
