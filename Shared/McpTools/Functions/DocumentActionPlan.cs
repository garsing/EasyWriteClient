using System;
using System.Collections.Generic;

namespace WordAddIn1
{
    /// <summary>
    /// process_actions 执行层规划：已按快照从后往前排好的动作 → Single / GeoRun。
    /// 无 COM、不调统一定位门面。
    /// </summary>
    public static class DocumentActionPlan
    {
        public sealed class Item
        {
            public int OriginalIndex { get; set; }
            public bool IsGeoCell { get; set; }
            public string TableId { get; set; }
        }

        public sealed class Run
        {
            public bool IsGeoRun { get; set; }
            public string TableId { get; set; }
            public List<int> OriginalIndices { get; set; } = new List<int>();
        }

        /// <summary>
        /// <paramref name="orderedDescending"/> 须已按文档位置从后往前排好。
        /// 相邻、同表、几何码写格（insert/replace/delete 可混）且 ≥2 条合并为 GeoRun。
        /// </summary>
        public static List<Run> BuildPlan(IList<Item> orderedDescending)
        {
            var steps = new List<Run>();
            if (orderedDescending == null || orderedDescending.Count == 0)
            {
                return steps;
            }

            int i = 0;
            while (i < orderedDescending.Count)
            {
                Item cur = orderedDescending[i];
                if (cur == null || !cur.IsGeoCell || string.IsNullOrEmpty(cur.TableId))
                {
                    steps.Add(Single(cur));
                    i++;
                    continue;
                }

                int j = i;
                while (j + 1 < orderedDescending.Count)
                {
                    Item next = orderedDescending[j + 1];
                    if (next == null
                        || !next.IsGeoCell
                        || !string.Equals(cur.TableId, next.TableId, StringComparison.Ordinal))
                    {
                        break;
                    }

                    j++;
                }

                if (j > i)
                {
                    var geo = new Run
                    {
                        IsGeoRun = true,
                        TableId = cur.TableId
                    };
                    for (int k = i; k <= j; k++)
                    {
                        geo.OriginalIndices.Add(orderedDescending[k].OriginalIndex);
                    }

                    steps.Add(geo);
                }
                else
                {
                    steps.Add(Single(cur));
                }

                i = j + 1;
            }

            return steps;
        }

        private static Run Single(Item item)
        {
            return new Run
            {
                IsGeoRun = false,
                TableId = item?.TableId,
                OriginalIndices = new List<int> { item == null ? 0 : item.OriginalIndex }
            };
        }
    }
}
