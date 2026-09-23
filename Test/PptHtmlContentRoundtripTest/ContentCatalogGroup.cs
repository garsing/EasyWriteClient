using System;
using System.Collections.Generic;
using WordAddIn1.PresentationHost;

namespace PptHtmlContentRoundtripTest
{
    internal static partial class ContentCatalog
    {
        private static void AddGroupCases(List<ContentCase> list, ref int page)
        {
            AddGroup(list, ref page, "grp-read-shell", s =>
            {
                if (!TrySeedGroup(s, ContentHtml.TwoTextboxes(), out string groupId, out string err))
                {
                    return err;
                }

                if (!s.TryReadPage(out PptHtmlReadResult pageRead, out err))
                {
                    return "整页读失败: " + err;
                }

                PptHtmlShapeNode group = ContentGroupSession.FindGroup(pageRead);
                if (group == null || group.ShapeId != groupId)
                {
                    return "顶层没有组 id=" + groupId;
                }

                if (ContentGroupSession.FindTextTop(pageRead, "源标题") != null
                    || ContentGroupSession.FindTextTop(pageRead, "源正文") != null)
                {
                    return "组内叶子被拆成顶层兄弟";
                }

                if (!ContentAssert.TryParseGeo(group.Style, out _, out _, out double w, out double h)
                    || w < 10 || h < 10)
                {
                    return "组没有页几何: " + group.Style;
                }

                return null;
            });

            AddGroup(list, ref page, "grp-read-expand-pct", s =>
            {
                if (!TrySeedGroup(s, ContentHtml.TwoTextboxes(), out string groupId, out string err))
                {
                    return err;
                }

                if (!s.TryReadPage(out PptHtmlReadResult pageRead, out err))
                {
                    return err;
                }

                PptHtmlShapeNode group = ContentGroupSession.FindGroup(pageRead)
                    ?? ContentGroupSession.FindByIdDeep(pageRead, groupId);
                if (!s.TryReadShape(groupId, out PptHtmlReadResult expand, out err))
                {
                    return "展开失败: " + err;
                }

                List<PptHtmlShapeNode> kids = ContentGroupSession.DirectChildren(expand, groupId);
                if (kids.Count < 2)
                {
                    return "展开子级数=" + kids.Count;
                }

                for (int i = 0; i < kids.Count; i++)
                {
                    PptHtmlShapeNode child = kids[i];
                    if (ContentGroupSession.IsGroup(child))
                    {
                        continue;
                    }

                    if (!ContentGroupSession.TopNearGroup(child, group))
                    {
                        return "孩子不是页%: " + child.Style + " 组=" + group.Style;
                    }
                }

                return null;
            });

            AddGroup(list, ref page, "grp-read-nested-shell", s =>
            {
                if (!TrySeedNested(s, out string outerId, out _, out string err))
                {
                    return err;
                }

                if (!s.TryReadShape(outerId, out PptHtmlReadResult expand, out err))
                {
                    return "展开外组失败: " + err;
                }

                PptHtmlShapeNode outer = ContentGroupSession.FindByIdDeep(expand, outerId)
                    ?? ContentGroupSession.FindGroup(expand);
                PptHtmlShapeNode inner = ContentGroupSession.FindChildGroup(outer);
                if (inner == null)
                {
                    return "外组展开不见内组";
                }

                if (string.IsNullOrEmpty(inner.ShapeId) || inner.ShapeId == outerId)
                {
                    return "内组没有自己的 ShapeId";
                }

                return ContentGroupSession.FindText(expand, "外组丙") == null
                    ? "外组展开缺外组丙"
                    : null;
            });

            AddGroup(list, ref page, "grp-read-leaf", s =>
            {
                if (!TrySeedGroup(s, ContentHtml.TwoTextboxes(), out string groupId, out string err))
                {
                    return err;
                }

                if (!s.TryReadShape(groupId, out PptHtmlReadResult expand, out err))
                {
                    return err;
                }

                List<string> leaves = ContentGroupSession.LeafIds(expand);
                if (leaves.Count == 0)
                {
                    return "展开无叶子";
                }

                if (!s.TryReadShape(leaves[0], out PptHtmlReadResult leaf, out err))
                {
                    return "读叶子失败: " + err;
                }

                PptHtmlShapeNode n = ContentGroupSession.FindByIdDeep(leaf, leaves[0]);
                if (n == null || ContentGroupSession.IsGroup(n))
                {
                    return "叶子详读不是 textbox";
                }

                return string.IsNullOrEmpty(n.Text) ? "叶子无正文" : null;
            });

            AddGroup(list, ref page, "grp-apply-lean-text", s =>
            {
                if (!TrySeedGroup(s, ContentHtml.TwoTextboxes(), out string groupId, out string err))
                {
                    return err;
                }

                if (!s.TryReadShape(groupId, out PptHtmlReadResult expand, out err))
                {
                    return err;
                }

                PptHtmlShapeNode title = ContentGroupSession.FindText(expand, "源标题");
                if (title == null)
                {
                    return "找不到源标题叶子";
                }

                string oldStyle = title.Style;
                if (!s.TryApply(ContentHtml.TextboxUpdate(title.ShapeId, "只改标题"), false, out _, out err))
                {
                    return "瘦稿失败: " + err;
                }

                if (!s.TryReadShape(groupId, out PptHtmlReadResult after, out err))
                {
                    return err;
                }

                if (ContentGroupSession.FindText(after, "只改标题") == null)
                {
                    return "正文没改";
                }

                if (ContentGroupSession.FindText(after, "源正文") == null)
                {
                    return "误伤另一叶子";
                }

                PptHtmlShapeNode changed = ContentGroupSession.FindText(after, "只改标题");
                if (!ContentAssert.GeoClose(oldStyle, changed.Style, 1.2))
                {
                    return "瘦稿改了几何: " + changed.Style;
                }

                return null;
            });

            AddGroup(list, ref page, "grp-neg-apply-group-text", s =>
            {
                if (!TrySeedGroup(s, ContentHtml.TwoTextboxes(), out string groupId, out string err))
                {
                    return err;
                }

                s.TryApply(ContentHtml.TextboxUpdate(groupId, "写到组上"), false, out _, out err);

                if (!s.TryReadShape(groupId, out PptHtmlReadResult expand, out string readErr))
                {
                    return readErr;
                }

                if (ContentGroupSession.FindText(expand, "源标题") == null
                    || ContentGroupSession.FindText(expand, "源正文") == null)
                {
                    return "对组改字伤到了叶子: " + err;
                }

                PptHtmlShapeNode leaked = ContentGroupSession.FindText(expand, "写到组上");
                if (leaked != null && !ContentGroupSession.IsGroup(leaked))
                {
                    return "把字写到叶子上了";
                }

                return null;
            });

            AddGroup(list, ref page, "grp-make-two", s =>
            {
                if (!TrySeedGroup(s, ContentHtml.TwoTextboxes(), out string groupId, out string err))
                {
                    return err;
                }

                if (!s.TryReadPage(out PptHtmlReadResult pageRead, out err))
                {
                    return err;
                }

                PptHtmlShapeNode group = ContentGroupSession.FindGroup(pageRead);
                if (group == null || group.ShapeId != groupId)
                {
                    return "组成组后无组";
                }

                if (!ContentAssert.TryParseGeo(group.Style, out _, out _, out double w, out double h)
                    || w < 20 || h < 20)
                {
                    return "组框不像包围盒: " + group.Style;
                }

                return null;
            });

            AddGroup(list, ref page, "grp-neg-make-one", s =>
            {
                if (!s.TryApply(ContentHtml.TwoTextboxes(), true, out List<string> ids, out string err)
                    || ids.Count < 1)
                {
                    return "铺底失败: " + err;
                }

                bool ok = s.TryGroup(new[] { ids[0] }, out _, out err);
                if (ok)
                {
                    return "1 个 id 组成组应失败";
                }

                if (string.IsNullOrEmpty(err) || err.IndexOf("至少 2", StringComparison.Ordinal) < 0)
                {
                    return "文案不对: " + err;
                }

                if (!s.TryReadPage(out PptHtmlReadResult pageRead, out string readErr))
                {
                    return readErr;
                }

                return ContentGroupSession.CountGroups(pageRead) == 0 ? null : "失败后出现了组";
            });

            AddGroup(list, ref page, "grp-neg-make-child", s =>
            {
                if (!TrySeedGroup(s, ContentHtml.TwoTextboxes(), out string groupId, out string err))
                {
                    return err;
                }

                if (!s.TryReadShape(groupId, out PptHtmlReadResult expand, out err))
                {
                    return err;
                }

                List<string> leaves = ContentGroupSession.LeafIds(expand);
                if (leaves.Count < 1)
                {
                    return "无叶子";
                }

                if (!s.TryApply(
                        ContentHtml.TextboxCreate("另顶层", left: 60, top: 20),
                        true,
                        out List<string> extra,
                        out err)
                    || extra.Count < 1)
                {
                    return "补顶层失败: " + err;
                }

                bool ok = s.TryGroup(new[] { leaves[0], extra[0] }, out _, out err);
                if (ok)
                {
                    return "组内叶子再 group 应失败";
                }

                return err != null && err.IndexOf("顶层", StringComparison.Ordinal) >= 0
                    ? null
                    : "文案不对: " + err;
            });

            AddGroup(list, ref page, "grp-neg-make-foreign", s =>
            {
                var args = new Dictionary<string, object>
                {
                    ["slide_id"] = "1",
                    ["shape_ids"] = new[] { "a", "b" },
                    ["left"] = 10
                };
                return ContentGroupSession.ToolRejects("group", args, "不接受")
                    ? null
                    : "group 带 left 应被工具层拒绝";
            });

            AddGroup(list, ref page, "grp-dup-same-page", s =>
            {
                if (!TrySeedGroup(s, ContentHtml.TwoTextboxes(), out string groupId, out string err))
                {
                    return err;
                }

                if (!s.TryReadPage(out PptHtmlReadResult before, out err))
                {
                    return err;
                }

                PptHtmlShapeNode src = ContentGroupSession.FindGroup(before);
                if (!s.TryDuplicate(groupId, 60, 10, out PresentationManageShapeResult dup, out err)
                    || string.IsNullOrEmpty(dup?.GroupShapeId))
                {
                    return "拷组失败: " + err;
                }

                if (!s.TryReadPage(out PptHtmlReadResult after, out err))
                {
                    return err;
                }

                if (ContentGroupSession.CountGroups(after) < 2)
                {
                    return "拷完组数=" + ContentGroupSession.CountGroups(after);
                }

                PptHtmlShapeNode neu = ContentGroupSession.FindByIdDeep(after, dup.GroupShapeId);
                if (neu == null)
                {
                    return "新组不在页上";
                }

                if (!ContentGroupSession.GeoNear(neu.Style, 60, 10, 2.5))
                {
                    return "新组位置: " + neu.Style;
                }

                if (!ContentAssert.TryParseGeo(src.Style, out _, out _, out double sw, out double sh)
                    || !ContentAssert.TryParseGeo(neu.Style, out _, out _, out double nw, out double nh)
                    || Math.Abs(sw - nw) > 2 || Math.Abs(sh - nh) > 2)
                {
                    return "尺寸变了 源=" + src.Style + " 新=" + neu.Style;
                }

                if (!s.TryReadShape(dup.GroupShapeId, out PptHtmlReadResult expand, out err))
                {
                    return err;
                }

                return ContentGroupSession.FindText(expand, "源标题") == null
                    ? "新组内文案丢了"
                    : null;
            });

            AddGroup(list, ref page, "grp-dup-chart", s =>
            {
                if (!TrySeedTitleAndPie(s, out string groupId, out string err))
                {
                    return err;
                }

                if (!s.TryReadPage(out PptHtmlReadResult before, out err))
                {
                    return err;
                }

                PptHtmlShapeNode src = ContentGroupSession.FindGroup(before);
                if (!s.TryDuplicate(groupId, 55, 12, out PresentationManageShapeResult dup, out err))
                {
                    return "拷组失败: " + err;
                }

                if (!s.TryReadShape(dup.GroupShapeId, out PptHtmlReadResult expand, out err))
                {
                    return err;
                }

                PptHtmlShapeNode chart = ContentGroupSession.FindByTypeDeep(expand, "chart");
                if (chart == null)
                {
                    return "新组内无 chart";
                }

                PptHtmlShapeNode neu = null;
                if (s.TryReadPage(out PptHtmlReadResult pageRead, out _))
                {
                    neu = ContentGroupSession.FindByIdDeep(pageRead, dup.GroupShapeId) ?? src;
                }

                if (!ContentGroupSession.TopNearGroup(chart, neu ?? src))
                {
                    return "新饼飞出组框: " + chart.Style + " 组=" + (neu == null ? "?" : neu.Style);
                }

                return null;
            });

            AddGroup(list, ref page, "grp-dup-nested", s =>
            {
                if (!TrySeedNested(s, out string outerId, out _, out string err))
                {
                    return err;
                }

                if (!s.TryDuplicate(outerId, 58, 8, out PresentationManageShapeResult dup, out err))
                {
                    return "拷外组失败: " + err;
                }

                if (!s.TryReadShape(dup.GroupShapeId, out PptHtmlReadResult expand, out err))
                {
                    return err;
                }

                PptHtmlShapeNode neu = ContentGroupSession.FindByIdDeep(expand, dup.GroupShapeId)
                    ?? ContentGroupSession.FindGroup(expand);
                return ContentGroupSession.FindChildGroup(neu) == null ? "新外组里没有内组" : null;
            });

            AddGroup(list, ref page, "grp-dup-inner", s =>
            {
                if (!TrySeedNested(s, out string outerId, out string innerId, out string err))
                {
                    return err;
                }

                if (!s.TryReadShape(outerId, out PptHtmlReadResult seeded, out err))
                {
                    return err;
                }

                PptHtmlShapeNode inner = ContentGroupSession.FindChildGroup(
                    ContentGroupSession.FindByIdDeep(seeded, outerId));
                if (inner == null || string.IsNullOrEmpty(inner.ShapeId))
                {
                    return "读回不见内组（seed=" + innerId + "）";
                }

                if (!s.TryDuplicate(inner.ShapeId, 62, 50, out PresentationManageShapeResult dup, out err))
                {
                    return "拷内组失败: " + err;
                }

                if (!s.TryReadPage(out PptHtmlReadResult pageRead, out err))
                {
                    return err;
                }

                if (ContentGroupSession.FindByIdDeep(pageRead, outerId) == null)
                {
                    return "外组被整份拷走了";
                }

                if (!s.TryReadShape(dup.GroupShapeId, out PptHtmlReadResult expand, out err))
                {
                    return err;
                }

                return ContentGroupSession.FindText(expand, "外组丙") != null
                    ? "拷内组带上了外组丙"
                    : null;
            });

            AddGroup(list, ref page, "grp-dup-lean-new", s =>
            {
                if (!TrySeedGroup(s, ContentHtml.TwoTextboxes(), out string groupId, out string err))
                {
                    return err;
                }

                if (!s.TryDuplicate(groupId, 58, 14, out PresentationManageShapeResult dup, out err))
                {
                    return err;
                }

                if (!s.TryReadShape(dup.GroupShapeId, out PptHtmlReadResult expand, out err))
                {
                    return err;
                }

                PptHtmlShapeNode title = ContentGroupSession.FindText(expand, "源标题");
                if (title == null)
                {
                    return "新组无源标题";
                }

                if (!s.TryApply(ContentHtml.TextboxUpdate(title.ShapeId, "只改新组"), false, out _, out err))
                {
                    return "改新叶子失败: " + err;
                }

                if (!s.TryReadShape(groupId, out PptHtmlReadResult src, out err))
                {
                    return err;
                }

                if (ContentGroupSession.FindText(src, "源标题") == null)
                {
                    return "源组正文被改了";
                }

                if (!s.TryReadShape(dup.GroupShapeId, out PptHtmlReadResult neu, out err))
                {
                    return err;
                }

                return ContentGroupSession.FindText(neu, "只改新组") == null ? "新组没改到" : null;
            });

            AddGroup(list, ref page, "grp-dup-cross-slide", s =>
            {
                if (!TrySeedGroup(s, ContentHtml.TwoTextboxes(), out string groupId, out string err))
                {
                    return err;
                }

                if (!s.TryAddBlankSlide(out string destSlide, out err))
                {
                    return "加页失败: " + err;
                }

                if (!s.TryDuplicate(groupId, 12, 16, out PresentationManageShapeResult dup, out err, destSlide))
                {
                    return "跨页拷失败: " + err;
                }

                string home = s.SlideId;
                s.SlideId = destSlide;
                bool destOk = s.TryReadPage(out PptHtmlReadResult destPage, out err)
                    && ContentGroupSession.FindByIdDeep(destPage, dup.GroupShapeId) != null;
                s.SlideId = home;
                if (!destOk)
                {
                    return "目标页没有新组: " + err;
                }

                if (!s.TryReadPage(out PptHtmlReadResult srcPage, out err))
                {
                    return err;
                }

                return ContentGroupSession.CountGroups(srcPage) >= 1 ? null : "源页组丢了";
            });

            AddGroup(list, ref page, "grp-dup-cross-channel", s =>
            {
                if (!TrySeedGroup(s, ContentHtml.TwoTextboxes(), out string groupId, out string err))
                {
                    return err;
                }

                if (!s.TryOpenSecondDeck(out object dest, out string destSlide, out err))
                {
                    return "SKIP: 开不了第二份稿: " + err;
                }

                try
                {
                    if (!s.TryDuplicate(
                            groupId, 10, 12, out PresentationManageShapeResult dup, out err, destSlide, dest))
                    {
                        return "跨稿拷失败: " + err;
                    }

                    return string.IsNullOrEmpty(dup?.GroupShapeId) ? "无新组 id" : null;
                }
                finally
                {
                    s.CloseDeckQuiet(dest);
                }
            });

            AddGroup(list, ref page, "grp-neg-dup-leaf", s =>
            {
                if (!s.TryApply(ContentHtml.TwoTextboxes(), true, out List<string> ids, out string err)
                    || ids.Count < 1)
                {
                    return err ?? "铺底失败";
                }

                bool ok = s.TryDuplicate(ids[0], 50, 10, out _, out err);
                if (ok)
                {
                    return "叶子拷组应失败";
                }

                return err != null && err.IndexOf("group", StringComparison.OrdinalIgnoreCase) >= 0
                    ? null
                    : "文案不对: " + err;
            });

            AddGroup(list, ref page, "grp-neg-dup-size", s =>
            {
                var args = new Dictionary<string, object>
                {
                    ["slide_id"] = "1",
                    ["shape_id"] = "g",
                    ["left"] = 10,
                    ["top"] = 10,
                    ["width"] = 40
                };
                return ContentGroupSession.ToolRejects("duplicate_group", args, "不接受")
                    ? null
                    : "带宽高应被工具层拒绝";
            });

            AddGroup(list, ref page, "grp-neg-dup-ids", s =>
            {
                var args = new Dictionary<string, object>
                {
                    ["slide_id"] = "1",
                    ["shape_id"] = "g",
                    ["left"] = 10,
                    ["top"] = 10,
                    ["shape_ids"] = new[] { "a" }
                };
                return ContentGroupSession.ToolRejects("duplicate_group", args, "shape_ids")
                    ? null
                    : "带 shape_ids 应被工具层拒绝";
            });

            AddGroup(list, ref page, "grp-neg-dup-cross-host", s =>
            {
                return "SKIP: ppt↔wpp 须两宿主，记手验";
            });
        }

        private static bool TrySeedGroup(
            ContentGroupSession s,
            string html,
            out string groupId,
            out string error)
        {
            groupId = null;
            if (!s.TryApply(html, true, out List<string> ids, out error) || ids.Count < 2)
            {
                error = "铺底 apply 失败: " + error + " ids=" + (ids == null ? 0 : ids.Count);
                return false;
            }

            if (!s.TryGroup(ids.ToArray(), out PresentationManageShapeResult grouped, out error)
                || string.IsNullOrEmpty(grouped?.GroupShapeId))
            {
                error = "组成组失败: " + error;
                return false;
            }

            groupId = grouped.GroupShapeId;
            return true;
        }

        private static bool TrySeedTitleAndPie(
            ContentGroupSession s,
            out string groupId,
            out string error)
        {
            groupId = null;
            if (!s.TryApply(ContentHtml.PieChart(), true, out List<string> chartIds, out error)
                || chartIds.Count < 1)
            {
                error = "SKIP: 铺饼失败: " + error;
                return false;
            }

            if (!s.TryApply(
                    ContentHtml.TextboxCreate("组内标题", left: 8, top: 12, width: 30, height: 10),
                    true,
                    out List<string> titleIds,
                    out error)
                || titleIds.Count < 1)
            {
                error = "铺标题失败: " + error;
                return false;
            }

            if (!s.TryGroup(new[] { titleIds[0], chartIds[0] }, out PresentationManageShapeResult grouped, out error)
                || string.IsNullOrEmpty(grouped?.GroupShapeId))
            {
                error = "标题+饼组成组失败: " + error;
                return false;
            }

            groupId = grouped.GroupShapeId;
            return true;
        }

        private static bool TrySeedNested(
            ContentGroupSession s,
            out string outerId,
            out string innerId,
            out string error)
        {
            outerId = null;
            innerId = null;
            if (!s.TryApply(ContentHtml.NestedGroups(), true, out _, out error))
            {
                error = "嵌套组 apply 失败: " + error;
                return false;
            }

            if (!s.TryReadPage(out PptHtmlReadResult pageRead, out error))
            {
                error = "嵌套组读页失败: " + error;
                return false;
            }

            PptHtmlShapeNode outer = ContentGroupSession.FindGroup(pageRead);
            if (outer == null || string.IsNullOrEmpty(outer.ShapeId))
            {
                error = "嵌套组 apply 后顶层无组";
                return false;
            }

            outerId = outer.ShapeId;
            if (!s.TryReadShape(outerId, out PptHtmlReadResult expand, out error))
            {
                error = "展开外组失败: " + error;
                return false;
            }

            PptHtmlShapeNode inner = ContentGroupSession.FindChildGroup(
                ContentGroupSession.FindByIdDeep(expand, outerId) ?? outer);
            if (inner == null || string.IsNullOrEmpty(inner.ShapeId))
            {
                error = "嵌套组 apply 后不见内组 树=" + ContentGroupSession.DescribeTree(expand);
                return false;
            }

            innerId = inner.ShapeId;
            return true;
        }

        private static void AddGroup(
            List<ContentCase> list,
            ref int page,
            string name,
            Func<ContentGroupSession, string> run)
        {
            page++;
            list.Add(new ContentCase
            {
                Index = list.Count + 1,
                Batch = 1,
                Page = page,
                Name = name,
                CreateHtml = a => ContentHtml.Section(""),
                GroupRun = run
            });
        }
    }
}
