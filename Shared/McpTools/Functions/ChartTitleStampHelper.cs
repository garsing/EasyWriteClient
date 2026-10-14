using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    /// <summary>
    /// 图表 Title 稳定编号：ew:C_ + 8 位。只负责发号/写戳，不按 Title 找图。
    /// 写失败不退回内容哈希。
    /// </summary>
    public static class ChartTitleStampHelper
    {
        public const int ChartTitleMaxChars = 255;
        public const string StampHead = "ew:";

        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        private static readonly Regex StampRegex = new Regex(
            @"^ew:C_([A-Za-z0-9]{8})",
            RegexOptions.Compiled);

        public static bool TryParse(string title, out string chartId, out string userSuffix)
        {
            chartId = null;
            userSuffix = "";
            if (string.IsNullOrEmpty(title))
            {
                return false;
            }

            Match m = StampRegex.Match(title);
            if (!m.Success)
            {
                return false;
            }

            chartId = "C_" + m.Groups[1].Value;
            string rest = title.Substring(m.Length);
            if (rest.StartsWith(" ", StringComparison.Ordinal))
            {
                rest = rest.Substring(1);
            }

            userSuffix = rest;
            return true;
        }

        public static string Format(string chartId, string userSuffix)
        {
            string stamp = StampHead + chartId;
            if (string.IsNullOrEmpty(userSuffix))
            {
                return stamp;
            }

            return stamp + " " + userSuffix;
        }

        public static bool TryFitToMaxLength(string formatted, out string fitted)
        {
            fitted = formatted ?? "";
            if (fitted.Length <= ChartTitleMaxChars)
            {
                return true;
            }

            Match m = StampRegex.Match(fitted);
            if (!m.Success)
            {
                return false;
            }

            string stamp = fitted.Substring(0, m.Length);
            bool hasSuffix = fitted.Length > m.Length && fitted[m.Length] == ' ';
            int minLen = hasSuffix ? stamp.Length + 1 : stamp.Length;
            if (minLen > ChartTitleMaxChars)
            {
                return false;
            }

            fitted = fitted.Substring(0, ChartTitleMaxChars);
            return true;
        }

        public static string NewRandomChartId(ISet<string> occupied)
        {
            if (occupied == null)
            {
                occupied = new HashSet<string>(StringComparer.Ordinal);
            }

            byte[] buf = new byte[8];
            for (int attempt = 0; attempt < 32; attempt++)
            {
                using (var rng = RandomNumberGenerator.Create())
                {
                    rng.GetBytes(buf);
                }

                char[] chars = new char[8];
                for (int i = 0; i < 8; i++)
                {
                    chars[i] = Alphabet[buf[i] % Alphabet.Length];
                }

                string id = "C_" + new string(chars);
                if (occupied.Add(id))
                {
                    return id;
                }
            }

            return null;
        }

        /// <summary>合法且未占用则沿用；否则新发。新发的号需要写回 Title。</summary>
        public static string DecideChartId(string title, ISet<string> occupied)
        {
            if (occupied == null)
            {
                throw new ArgumentNullException(nameof(occupied));
            }

            if (TryParse(title, out string existing, out _) && occupied.Add(existing))
            {
                return existing;
            }

            return NewRandomChartId(occupied);
        }

        public static bool TryWriteStamp(ChartResolveHelper.BodyChart chart, string chartId)
        {
            if (chart == null || string.IsNullOrEmpty(chartId) || !chartId.StartsWith("C_", StringComparison.Ordinal))
            {
                return false;
            }

            string suffix = "";
            if (chart.TryReadTitle(out string current) && !string.IsNullOrEmpty(current))
            {
                if (TryParse(current, out _, out string parsedSuffix))
                {
                    suffix = parsedSuffix;
                }
                else
                {
                    suffix = current;
                }
            }

            string formatted = Format(chartId, suffix);
            if (!TryFitToMaxLength(formatted, out formatted))
            {
                return false;
            }

            return chart.TryWriteTitle(formatted);
        }

        /// <summary>
        /// 按序把编号写入形状 Title。已是该号则跳过。任一写失败返回 false，不改发哈希号。
        /// </summary>
        public static bool ApplyAssignedIds(
            IList<ChartResolveHelper.BodyChart> charts,
            IList<string> assignedIds)
        {
            if (assignedIds == null || assignedIds.Count == 0)
            {
                return true;
            }

            if (charts == null || charts.Count != assignedIds.Count)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[ChartStamp] 图数与编号数不一致，拒绝写入");
                return false;
            }

            for (int i = 0; i < assignedIds.Count; i++)
            {
                string wantId = assignedIds[i];
                if (string.IsNullOrEmpty(wantId))
                {
                    return false;
                }

                ChartResolveHelper.BodyChart chart = charts[i];
                bool alreadyOk = chart.TryReadTitle(out string current)
                    && TryParse(current, out string existingId, out _)
                    && existingId == wantId;
                if (alreadyOk)
                {
                    continue;
                }

                if (!TryWriteStamp(chart, wantId))
                {
                    System.Diagnostics.Debug.WriteLine($"[ChartStamp] i={i} 写 Title 失败 {wantId}");
                    return false;
                }

                System.Diagnostics.Debug.WriteLine($"[ChartStamp] i={i} 写入 {wantId}");
            }

            return true;
        }
    }
}
