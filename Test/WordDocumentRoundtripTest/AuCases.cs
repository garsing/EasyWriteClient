using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>自动编号只作读文标记。对照 docs/测试/word测试/自动编号只作读文标记。前缀 au-。</summary>
    internal static class AuCases
    {
        public const int Target = 37;

        private static readonly Regex AutonumTag = new Regex(
            @"<autonum\b[^>]*\bv\s*=\s*""([^""]*)""[^>]*/>",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static HashSet<string> Only;

        public static void Run(TestRun run, string host, string filter)
        {
            Only = Pick(filter, "au");
            if (Only != null)
            {
                run.TotalTarget = Only.Count;
            }

            var session = new WdSession(host);
            Read(run, session);
            Write(run, session);
            Clear(run, session);
            More(run, session);
        }

        private static void Read(TestRun run, WdSession session)
        {
            Case(run, "au-read-decimal", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r乙\r"))
                {
                    if (!doc.Ready || !session.ApplyNumberDefault(doc.Path, 1, 2))
                    {
                        return "铺编号失败 " + doc.Id;
                    }

                    if (!doc.ReadParagraphs())
                    {
                        return "读段失败 " + doc.Snapshot();
                    }

                    List<string> lists = session.ParaListStrings(doc.Path);
                    List<string> stories = session.ParaStories(doc.Path);
                    string jia = Inner(doc, "甲");
                    string yi = Inner(doc, "乙");
                    return TagThen(jia, "甲", At(lists, 0))
                        && TagThen(yi, "乙", At(lists, 1))
                        && Story(At(stories, 0)) == "甲"
                        && Story(At(stories, 1)) == "乙"
                        && doc.Snapshot().IndexOf("[list]", StringComparison.OrdinalIgnoreCase) < 0
                        ? null
                        : "甲=" + jia + " 乙=" + yi + " lists=" + string.Join("|", lists.ToArray());
                }
            });

            Case(run, "au-read-bullet", () =>
            {
                using (FormatProbe doc = Open(session, "项\r"))
                {
                    if (!doc.Ready || !session.ApplyBulletDefault(doc.Path, 1, 1))
                    {
                        return "铺项目符号失败";
                    }

                    if (!doc.ReadParagraphs())
                    {
                        return "读段失败";
                    }

                    string inner = Inner(doc, "项");
                    string list = At(session.ParaListStrings(doc.Path), 0);
                    return TagThen(inner, "项", list) && Story(At(session.ParaStories(doc.Path), 0)) == "项"
                        ? null
                        : inner + " list=" + list;
                }
            });

            Case(run, "au-read-hand", () =>
            {
                using (FormatProbe doc = Open(session, "1. 不是列表\r"))
                {
                    if (!doc.Ready || !doc.ReadParagraphs())
                    {
                        return "读段失败";
                    }

                    string inner = Inner(doc, "不是列表");
                    string list = At(session.ParaListStrings(doc.Path), 0).Trim();
                    string story = Story(At(session.ParaStories(doc.Path), 0));
                    return inner.IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) < 0
                        && inner.IndexOf("1. 不是列表", StringComparison.Ordinal) >= 0
                        && story.IndexOf("1.", StringComparison.Ordinal) >= 0
                        && list.Length == 0
                        ? null
                        : inner + " story=" + story + " list=" + list;
                }
            });

            Case(run, "au-read-para", () =>
            {
                using (FormatProbe doc = Open(session, "第一条\r"))
                {
                    if (!doc.Ready || !session.ApplyNumberDefault(doc.Path, 1, 1) || !doc.ReadParagraphs())
                    {
                        return "读段失败";
                    }

                    string inner = ParaInner(doc, "第一条");
                    string list = At(session.ParaListStrings(doc.Path), 0);
                    return TagThen(inner, "第一条", list) ? null : inner + " list=" + list;
                }
            });

            Case(run, "au-read-two-sent", () =>
            {
                using (FormatProbe doc = Open(session, "甲。乙。\r"))
                {
                    if (!doc.Ready || !session.ApplyNumberDefault(doc.Path, 1, 1) || !doc.ReadParagraphs())
                    {
                        return "读段失败 " + doc.Snapshot();
                    }

                    string first = doc.Paragraphs.Count > 0 ? (doc.Paragraphs[0].Text ?? "") : "";
                    int tags = AutonumTag.Matches(first).Count;
                    return tags == 1
                        && first.IndexOf("甲", StringComparison.Ordinal) >= 0
                        && first.IndexOf("乙", StringComparison.Ordinal) >= 0
                        && first.IndexOf("甲", StringComparison.Ordinal) < first.IndexOf("乙", StringComparison.Ordinal)
                        ? null
                        : first;
                }
            });

            Case(run, "au-read-beside", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r乙\r"))
                {
                    if (!doc.Ready || !session.ApplyNumberDefault(doc.Path, 1, 1) || !doc.ReadParagraphs())
                    {
                        return "读段失败";
                    }

                    string jia = Inner(doc, "甲");
                    string yi = Inner(doc, "乙");
                    return jia.IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) >= 0
                        && yi.IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) < 0
                        ? null
                        : "甲=" + jia + " 乙=" + yi;
                }
            });

            Case(run, "au-read-table", () =>
            {
                using (FormatProbe doc = Open(session, "外\r"))
                {
                    if (!doc.Ready
                        || !session.InsertTableBeforePara(doc.Path, 1, "格甲")
                        || !session.NumberFirstTableCell(doc.Path)
                        || !doc.ReadParagraphs())
                    {
                        return "铺表失败 " + doc.Snapshot();
                    }

                    string cell = Inner(doc, "格甲");
                    string outside = Inner(doc, "外");
                    return cell.IndexOf("<table", StringComparison.OrdinalIgnoreCase) < 0
                        && cell.IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) >= 0
                        && cell.IndexOf("格甲", StringComparison.Ordinal) >= 0
                        && outside.IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) < 0
                        ? null
                        : "格=" + cell + " 外=" + outside + " " + doc.Snapshot();
                }
            });

            Case(run, "au-read-empty", () =>
            {
                using (FormatProbe doc = Open(session, "占位\r"))
                {
                    if (!doc.Ready
                        || !session.ApplyNumberDefault(doc.Path, 1, 1)
                        || !session.ClearParaKeepList(doc.Path, 1))
                    {
                        return "铺空编号段失败";
                    }

                    string body = doc.Snapshot();
                    string list = At(session.ParaListStrings(doc.Path), 0).Trim();
                    string story = Story(At(session.ParaStories(doc.Path), 0));
                    return body.IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) >= 0
                        && story.Length == 0
                        && list.Length > 0
                        ? null
                        : "story=[" + story + "] list=[" + list + "] " + body;
                }
            });
        }

        private static void Write(TestRun run, WdSession session)
        {
            Case(run, "au-rep-body", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 1))
                    {
                        return "铺失败";
                    }

                    ToolResult written = session.ProcessActions(doc.Id, Replace(doc.Paragraph("甲"), "新正文"));
                    if (written == null || !written.Success || !doc.ReadParagraphs())
                    {
                        return written == null ? "无回包" : (written.Error ?? doc.Snapshot());
                    }

                    string list = At(session.ParaListStrings(doc.Path), 0);
                    string story = Story(At(session.ParaStories(doc.Path), 0));
                    string inner = Inner(doc, "新正文");
                    return story == "新正文" && list.Trim().Length > 0 && TagThen(inner, "新正文", list)
                        ? null
                        : "story=" + story + " list=" + list + " inner=" + inner;
                }
            });

            Case(run, "au-rep-copy", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 1))
                    {
                        return "铺失败";
                    }

                    ToolResult written = session.ProcessActions(doc.Id, Replace(doc.Paragraph("甲"), "1. 抄来的"));
                    if (written == null || !written.Success)
                    {
                        return written == null ? "无回包" : written.Error;
                    }

                    string story = Story(At(session.ParaStories(doc.Path), 0));
                    string list = At(session.ParaListStrings(doc.Path), 0).Trim();
                    return story.IndexOf("1. 抄来的", StringComparison.Ordinal) >= 0
                        && story != "抄来的"
                        && list.Length > 0
                        ? null
                        : "story=" + story + " list=" + list;
                }
            });

            Case(run, "au-rep-two", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r乙\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 2))
                    {
                        return "铺失败";
                    }

                    string codes = doc.Paragraph("甲") + "," + doc.Paragraph("乙");
                    ToolResult written = session.ProcessActions(doc.Id, Replace(codes, "丙\r丁"));
                    if (written == null || !written.Success)
                    {
                        return written == null ? "无回包" : written.Error;
                    }

                    List<string> stories = session.ParaStories(doc.Path);
                    List<string> lists = session.ParaListStrings(doc.Path);
                    return Story(At(stories, 0)) == "丙"
                        && Story(At(stories, 1)) == "丁"
                        && At(lists, 0).Trim().Length > 0
                        && At(lists, 1).Trim().Length > 0
                        ? null
                        : string.Join("|", stories.ToArray()) + " / " + string.Join("|", lists.ToArray());
                }
            });

            Case(run, "au-rep-hand", () =>
            {
                using (FormatProbe doc = Open(session, "1. 不是列表\r"))
                {
                    if (!doc.Ready || !doc.ReadParagraphs())
                    {
                        return "读段失败";
                    }

                    ToolResult written = session.ProcessActions(doc.Id, Replace(doc.Paragraph("不是列表"), "1. 改了"));
                    if (written == null || !written.Success || !doc.ReadParagraphs())
                    {
                        return written == null ? "无回包" : (written.Error ?? "重读失败");
                    }

                    string story = Story(At(session.ParaStories(doc.Path), 0));
                    string list = At(session.ParaListStrings(doc.Path), 0).Trim();
                    string inner = Inner(doc, "改了");
                    return story.IndexOf("1. 改了", StringComparison.Ordinal) >= 0
                        && list.Length == 0
                        && inner.IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) < 0
                        ? null
                        : "story=" + story + " list=" + list + " inner=" + inner;
                }
            });

            Case(run, "au-ins-plain", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r乙\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 2))
                    {
                        return "铺失败";
                    }

                    ToolResult written = session.ProcessActions(doc.Id, InsertAfter(doc.Paragraph("甲"), "hello"));
                    if (written == null || !written.Success)
                    {
                        return written == null ? "无回包" : written.Error;
                    }

                    string blob = string.Join("\n", session.ParaStories(doc.Path).ToArray());
                    string list = At(session.ParaListStrings(doc.Path), 0).Trim();
                    return blob.IndexOf("hello", StringComparison.Ordinal) >= 0 && list.Length > 0
                        ? null
                        : blob + " list=" + list;
                }
            });

            Case(run, "au-ins-copy", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 1))
                    {
                        return "铺失败";
                    }

                    ToolResult written = session.ProcessActions(doc.Id, InsertAfter(doc.Paragraph("甲"), "2. 插上"));
                    if (written == null || !written.Success)
                    {
                        return written == null ? "无回包" : written.Error;
                    }

                    string blob = string.Join("\n", session.ParaStories(doc.Path).ToArray());
                    return blob.IndexOf("2. 插上", StringComparison.Ordinal) >= 0 ? null : blob;
                }
            });

            Case(run, "au-ins-tag", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 1))
                    {
                        return "铺失败";
                    }

                    const string literal = "<autonum v=\"1.\"/>正文";
                    ToolResult written = session.ProcessActions(doc.Id, InsertAfter(doc.Paragraph("甲"), literal));
                    if (written == null || !written.Success)
                    {
                        return written == null ? "无回包" : written.Error;
                    }

                    string blob = string.Join("\n", session.ParaStories(doc.Path).ToArray());
                    return blob.IndexOf(literal, StringComparison.Ordinal) >= 0 ? null : blob;
                }
            });
        }

        private static void Clear(TestRun run, WdSession session)
        {
            Case(run, "au-clear-number", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r乙\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 2) || !doc.ReadParagraphs())
                    {
                        return "铺失败";
                    }

                    ToolResult cleared = session.ApplyParagraphFormat(new Dictionary<string, object>
                    {
                        ["channel_id"] = doc.Id,
                        ["target_paragraph_codes"] = doc.Paragraph("甲"),
                        ["para_format"] = new Dictionary<string, object>
                        {
                            ["list"] = new Dictionary<string, object> { ["kind"] = "none" }
                        }
                    });
                    if (cleared == null || !cleared.Success || !doc.ReadParagraphs())
                    {
                        return cleared == null ? "无回包" : (cleared.Error ?? "重读失败");
                    }

                    string jiaList = At(session.ParaListStrings(doc.Path), 0).Trim();
                    string jia = Inner(doc, "甲");
                    string yi = Inner(doc, "乙");
                    return jiaList.Length == 0
                        && jia.IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) < 0
                        && yi.IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) >= 0
                        ? null
                        : "甲=" + jia + " 乙=" + yi + " list=" + jiaList;
                }
            });

            Case(run, "au-clear-bullet", () =>
            {
                using (FormatProbe doc = Open(session, "项\r"))
                {
                    if (!doc.Ready || !session.ApplyBulletDefault(doc.Path, 1, 1) || !doc.ReadParagraphs())
                    {
                        return "铺失败";
                    }

                    ToolResult cleared = session.ApplyParagraphFormat(new Dictionary<string, object>
                    {
                        ["channel_id"] = doc.Id,
                        ["target_paragraph_codes"] = doc.Paragraph("项"),
                        ["para_format"] = new Dictionary<string, object>
                        {
                            ["list"] = new Dictionary<string, object> { ["kind"] = "none" }
                        }
                    });
                    if (cleared == null || !cleared.Success || !doc.ReadParagraphs())
                    {
                        return cleared == null ? "无回包" : (cleared.Error ?? "重读失败");
                    }

                    string list = At(session.ParaListStrings(doc.Path), 0).Trim();
                    string inner = Inner(doc, "项");
                    string story = Story(At(session.ParaStories(doc.Path), 0));
                    return list.Length == 0
                        && inner.IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) < 0
                        && story == "项"
                        ? null
                        : "story=" + story + " list=" + list + " inner=" + inner;
                }
            });
        }

        private static void More(TestRun run, WdSession session)
        {
            Case(run, "au-read-three", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r乙\r丙\r"))
                {
                    if (!doc.Ready || !session.ApplyNumberDefault(doc.Path, 1, 3) || !doc.ReadParagraphs())
                    {
                        return "铺失败 " + doc.Snapshot();
                    }

                    List<string> lists = session.ParaListStrings(doc.Path);
                    List<string> stories = session.ParaStories(doc.Path);
                    return TaggedBody(doc, "甲", ListFor(stories, lists, "甲"))
                        && TaggedBody(doc, "乙", ListFor(stories, lists, "乙"))
                        && TaggedBody(doc, "丙", ListFor(stories, lists, "丙"))
                        && Story(At(stories, 0)) == "甲"
                        && Story(At(stories, 1)) == "乙"
                        && Story(At(stories, 2)) == "丙"
                        && doc.Snapshot().IndexOf("[list]", StringComparison.OrdinalIgnoreCase) < 0
                        ? null
                        : doc.Snapshot();
                }
            });

            Case(run, "au-read-same", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r甲\r"))
                {
                    if (!doc.Ready || !session.ApplyNumberDefault(doc.Path, 1, 2) || !doc.ReadParagraphs())
                    {
                        return "铺失败 " + doc.Snapshot();
                    }

                    List<string> lists = session.ParaListStrings(doc.Path);
                    int tagged = 0;
                    for (int i = 0; i < doc.Paragraphs.Count; i++)
                    {
                        string text = doc.Paragraphs[i].Text ?? "";
                        if (text.IndexOf("甲", StringComparison.Ordinal) < 0)
                        {
                            continue;
                        }

                        if (!TagThen(text, "甲", At(lists, i)))
                        {
                            return text + " list=" + At(lists, i);
                        }

                        tagged++;
                    }

                    return tagged == 2 ? null : "tagged=" + tagged + " " + doc.Snapshot();
                }
            });

            Case(run, "au-read-mid", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r乙\r丙\r"))
                {
                    if (!doc.Ready || !NumberBody(session, doc.Path, "乙") || !doc.ReadParagraphs())
                    {
                        return "铺失败 " + doc.Snapshot();
                    }

                    string yi = ListFor(session.ParaStories(doc.Path), session.ParaListStrings(doc.Path), "乙");
                    return TaggedBody(doc, "乙", yi)
                        && NoTag(Inner(doc, "甲"))
                        && NoTag(Inner(doc, "丙"))
                        ? null
                        : doc.Snapshot();
                }
            });

            Case(run, "au-read-hand-cn", () =>
            {
                using (FormatProbe doc = Open(session, "一、不是列表\r"))
                {
                    if (!doc.Ready || !doc.ReadParagraphs())
                    {
                        return "读段失败";
                    }

                    string inner = Inner(doc, "不是列表");
                    string story = Story(At(session.ParaStories(doc.Path), 0));
                    string list = At(session.ParaListStrings(doc.Path), 0).Trim();
                    return NoTag(inner)
                        && inner.IndexOf("一、不是列表", StringComparison.Ordinal) >= 0
                        && story.IndexOf("一、", StringComparison.Ordinal) >= 0
                        && list.Length == 0
                        ? null
                        : inner + " story=" + story + " list=" + list;
                }
            });

            Case(run, "au-read-bullets", () =>
            {
                using (FormatProbe doc = Open(session, "项甲\r项乙\r"))
                {
                    if (!doc.Ready || !session.ApplyBulletDefault(doc.Path, 1, 2) || !doc.ReadParagraphs())
                    {
                        return "铺失败 " + doc.Snapshot();
                    }

                    List<string> lists = session.ParaListStrings(doc.Path);
                    List<string> stories = session.ParaStories(doc.Path);
                    return TaggedBody(doc, "项甲", ListFor(stories, lists, "项甲"))
                        && TaggedBody(doc, "项乙", ListFor(stories, lists, "项乙"))
                        && Story(At(stories, 0)) == "项甲"
                        && Story(At(stories, 1)) == "项乙"
                        ? null
                        : doc.Snapshot();
                }
            });

            Case(run, "au-read-mix", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r项\r"))
                {
                    if (!doc.Ready
                        || !session.ApplyNumberDefault(doc.Path, 1, 1)
                        || !session.ApplyBulletDefault(doc.Path, 2, 2)
                        || !doc.ReadParagraphs())
                    {
                        return "铺失败 " + doc.Snapshot();
                    }

                    List<string> lists = session.ParaListStrings(doc.Path);
                    List<string> stories = session.ParaStories(doc.Path);
                    return TaggedBody(doc, "甲", ListFor(stories, lists, "甲"))
                        && TaggedBody(doc, "项", ListFor(stories, lists, "项"))
                        ? null
                        : doc.Snapshot();
                }
            });

            Case(run, "au-read-long", () =>
            {
                string body = new string('测', 40);
                using (FormatProbe doc = Open(session, body + "\r"))
                {
                    if (!doc.Ready || !session.ApplyNumberDefault(doc.Path, 1, 1) || !doc.ReadParagraphs())
                    {
                        return "铺失败";
                    }

                    string inner = Inner(doc, "测");
                    string list = At(session.ParaListStrings(doc.Path), 0);
                    return AutonumTag.Matches(inner).Count == 1
                        && TagThen(inner, body, list)
                        && Story(At(session.ParaStories(doc.Path), 0)) == body
                        ? null
                        : inner;
                }
            });

            Case(run, "au-read-symbol", () =>
            {
                const string body = "甲<乙>&丙";
                using (FormatProbe doc = Open(session, body + "\r"))
                {
                    if (!doc.Ready || !session.ApplyNumberDefault(doc.Path, 1, 1) || !doc.ReadParagraphs())
                    {
                        return "铺失败 " + doc.Snapshot();
                    }

                    string inner = Inner(doc, "甲");
                    string list = At(session.ParaListStrings(doc.Path), 0);
                    return TagThen(inner, body, list)
                        && Story(At(session.ParaStories(doc.Path), 0)) == body
                        ? null
                        : "story=" + Story(At(session.ParaStories(doc.Path), 0)) + " inner=" + inner;
                }
            });

            Case(run, "au-read-gap", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r\r乙\r"))
                {
                    if (!doc.Ready
                        || !NumberBody(session, doc.Path, "甲")
                        || !NumberBody(session, doc.Path, "乙")
                        || !doc.ReadParagraphs())
                    {
                        return "铺失败 " + doc.Snapshot();
                    }

                    List<string> lists = session.ParaListStrings(doc.Path);
                    List<string> stories = session.ParaStories(doc.Path);
                    int tags = 0;
                    foreach (Marker marker in doc.Paragraphs)
                    {
                        if ((marker.Text ?? "").IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            tags++;
                        }
                    }

                    return tags == 2
                        && TaggedBody(doc, "甲", ListFor(stories, lists, "甲"))
                        && TaggedBody(doc, "乙", ListFor(stories, lists, "乙"))
                        ? null
                        : "tags=" + tags + " " + doc.Snapshot();
                }
            });

            Case(run, "au-read-cell-out", () =>
            {
                using (FormatProbe doc = Open(session, "外\r"))
                {
                    if (!doc.Ready
                        || !session.InsertTableBeforePara(doc.Path, 1, "格甲")
                        || !session.NumberFirstTableCell(doc.Path)
                        || !NumberBody(session, doc.Path, "外")
                        || !doc.ReadParagraphs())
                    {
                        return "铺失败 " + doc.Snapshot();
                    }

                    List<string> lists = session.ParaListStrings(doc.Path);
                    List<string> stories = session.ParaStories(doc.Path);
                    string cell = Inner(doc, "格甲");
                    return cell.IndexOf("<table", StringComparison.OrdinalIgnoreCase) < 0
                        && TaggedBody(doc, "格甲", ListFor(stories, lists, "格甲"))
                        && TaggedBody(doc, "外", ListFor(stories, lists, "外"))
                        ? null
                        : doc.Snapshot();
                }
            });

            Case(run, "au-rep-tag", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 1))
                    {
                        return "铺失败";
                    }

                    const string literal = "<autonum v=\"9.\"/>正文";
                    ToolResult written = session.ProcessActions(doc.Id, Replace(doc.Paragraph("甲"), literal));
                    if (written == null || !written.Success || !doc.ReadParagraphs())
                    {
                        return written == null ? "无回包" : (written.Error ?? doc.Snapshot());
                    }

                    string story = Story(At(session.ParaStories(doc.Path), 0));
                    string list = At(session.ParaListStrings(doc.Path), 0);
                    return story.IndexOf(literal, StringComparison.Ordinal) >= 0
                        && list.Trim().Length > 0
                        && TagThen(Inner(doc, "正文"), literal, list)
                        ? null
                        : "story=" + story + " inner=" + Inner(doc, "正文");
                }
            });

            Case(run, "au-rep-cn", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 1))
                    {
                        return "铺失败";
                    }

                    ToolResult written = session.ProcessActions(doc.Id, Replace(doc.Paragraph("甲"), "一、正文"));
                    if (written == null || !written.Success || !doc.ReadParagraphs())
                    {
                        return written == null ? "无回包" : (written.Error ?? "重读失败");
                    }

                    string list = At(session.ParaListStrings(doc.Path), 0);
                    string story = Story(At(session.ParaStories(doc.Path), 0));
                    return story.IndexOf("一、正文", StringComparison.Ordinal) >= 0
                        && list.Trim().Length > 0
                        && TagThen(Inner(doc, "正文"), "一、正文", list)
                        ? null
                        : "story=" + story + " inner=" + Inner(doc, "正文");
                }
            });

            Case(run, "au-rep-symbol", () =>
            {
                const string body = "新<文>&字";
                using (FormatProbe doc = Open(session, "甲\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 1))
                    {
                        return "铺失败";
                    }

                    ToolResult written = session.ProcessActions(doc.Id, Replace(doc.Paragraph("甲"), body));
                    if (written == null || !written.Success || !doc.ReadParagraphs())
                    {
                        return written == null ? "无回包" : (written.Error ?? doc.Snapshot());
                    }

                    string list = At(session.ParaListStrings(doc.Path), 0);
                    return Story(At(session.ParaStories(doc.Path), 0)) == body
                        && list.Trim().Length > 0
                        && TagThen(Inner(doc, "新"), body, list)
                        ? null
                        : "story=" + Story(At(session.ParaStories(doc.Path), 0)) + " inner=" + Inner(doc, "新");
                }
            });

            Case(run, "au-rep-mid", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r乙\r丙\r"))
                {
                    if (!doc.Ready || !session.ApplyNumberDefault(doc.Path, 1, 3) || !doc.ReadParagraphs())
                    {
                        return "铺失败";
                    }

                    ToolResult written = session.ProcessActions(doc.Id, Replace(doc.Paragraph("乙"), "改中"));
                    if (written == null || !written.Success || !doc.ReadParagraphs())
                    {
                        return written == null ? "无回包" : (written.Error ?? doc.Snapshot());
                    }

                    List<string> lists = session.ParaListStrings(doc.Path);
                    List<string> stories = session.ParaStories(doc.Path);
                    return Story(At(stories, 1)) == "改中"
                        && At(lists, 1).Trim().Length > 0
                        && TaggedBody(doc, "甲", ListFor(stories, lists, "甲"))
                        && TaggedBody(doc, "改中", ListFor(stories, lists, "改中"))
                        && TaggedBody(doc, "丙", ListFor(stories, lists, "丙"))
                        ? null
                        : doc.Snapshot();
                }
            });

            Case(run, "au-ins-start", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 1))
                    {
                        return "铺失败";
                    }

                    ToolResult written = session.ProcessActions(doc.Id, InsertStart("前言\r"));
                    if (written == null || !written.Success || !doc.ReadParagraphs())
                    {
                        return written == null ? "无回包" : (written.Error ?? doc.Snapshot());
                    }

                    List<string> lists = session.ParaListStrings(doc.Path);
                    List<string> stories = session.ParaStories(doc.Path);
                    string plain = ListFor(stories, lists, "前言");
                    bool plainOk = plain.Trim().Length > 0
                        ? TaggedBody(doc, "前言", plain)
                        : NoTag(Inner(doc, "前言"));
                    return plainOk
                        && TaggedBody(doc, "甲", ListFor(stories, lists, "甲"))
                        && StoryHas(stories, "前言")
                        && StoryHas(stories, "甲")
                        ? null
                        : doc.Snapshot();
                }
            });

            Case(run, "au-ins-multi", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r乙\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 2))
                    {
                        return "铺失败";
                    }

                    ToolResult written = session.ProcessActions(doc.Id, InsertAfter(doc.Paragraph("甲"), "丙\r丁"));
                    if (written == null || !written.Success)
                    {
                        return written == null ? "无回包" : written.Error;
                    }

                    List<string> stories = session.ParaStories(doc.Path);
                    string list = At(session.ParaListStrings(doc.Path), 0).Trim();
                    return list.Length > 0
                        && StoryHas(stories, "丙")
                        && StoryHas(stories, "丁")
                        && StoryHas(stories, "甲")
                        ? null
                        : string.Join("|", stories.ToArray());
                }
            });

            Case(run, "au-clear-mid", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r乙\r丙\r"))
                {
                    if (!doc.Ready || !session.ApplyNumberDefault(doc.Path, 1, 3) || !doc.ReadParagraphs())
                    {
                        return "铺失败";
                    }

                    ToolResult cleared = ClearCodes(session, doc, doc.Paragraph("乙"));
                    if (cleared == null || !cleared.Success || !doc.ReadParagraphs())
                    {
                        return cleared == null ? "无回包" : (cleared.Error ?? "重读失败");
                    }

                    List<string> lists = session.ParaListStrings(doc.Path);
                    List<string> stories = session.ParaStories(doc.Path);
                    return ListFor(stories, lists, "乙").Trim().Length == 0
                        && NoTag(Inner(doc, "乙"))
                        && TaggedBody(doc, "甲", ListFor(stories, lists, "甲"))
                        && TaggedBody(doc, "丙", ListFor(stories, lists, "丙"))
                        ? null
                        : doc.Snapshot();
                }
            });

            Case(run, "au-clear-both", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r乙\r"))
                {
                    if (!ReadyNumbered(session, doc, 1, 2))
                    {
                        return "铺失败";
                    }

                    string codes = doc.Paragraph("甲") + "," + doc.Paragraph("乙");
                    ToolResult cleared = ClearCodes(session, doc, codes);
                    if (cleared == null || !cleared.Success || !doc.ReadParagraphs())
                    {
                        return cleared == null ? "无回包" : (cleared.Error ?? "重读失败");
                    }

                    return NoTag(Inner(doc, "甲"))
                        && NoTag(Inner(doc, "乙"))
                        && At(session.ParaListStrings(doc.Path), 0).Trim().Length == 0
                        && At(session.ParaListStrings(doc.Path), 1).Trim().Length == 0
                        && Story(At(session.ParaStories(doc.Path), 0)) == "甲"
                        && Story(At(session.ParaStories(doc.Path), 1)) == "乙"
                        ? null
                        : doc.Snapshot();
                }
            });

            Case(run, "au-clear-last", () =>
            {
                using (FormatProbe doc = Open(session, "甲\r乙\r丙\r"))
                {
                    if (!doc.Ready || !session.ApplyNumberDefault(doc.Path, 1, 3) || !doc.ReadParagraphs())
                    {
                        return "铺失败";
                    }

                    ToolResult cleared = ClearCodes(session, doc, doc.Paragraph("丙"));
                    if (cleared == null || !cleared.Success || !doc.ReadParagraphs())
                    {
                        return cleared == null ? "无回包" : (cleared.Error ?? "重读失败");
                    }

                    List<string> lists = session.ParaListStrings(doc.Path);
                    List<string> stories = session.ParaStories(doc.Path);
                    return ListFor(stories, lists, "丙").Trim().Length == 0
                        && NoTag(Inner(doc, "丙"))
                        && TaggedBody(doc, "甲", ListFor(stories, lists, "甲"))
                        && TaggedBody(doc, "乙", ListFor(stories, lists, "乙"))
                        ? null
                        : doc.Snapshot();
                }
            });

            Case(run, "au-clear-plain", () =>
            {
                using (FormatProbe doc = Open(session, "普通\r"))
                {
                    if (!doc.Ready || !doc.ReadParagraphs())
                    {
                        return "读段失败";
                    }

                    ToolResult cleared = ClearCodes(session, doc, doc.Paragraph("普通"));
                    if (cleared == null || !cleared.Success || !doc.ReadParagraphs())
                    {
                        return cleared == null ? "无回包" : (cleared.Error ?? "重读失败");
                    }

                    return NoTag(Inner(doc, "普通"))
                        && Story(At(session.ParaStories(doc.Path), 0)) == "普通"
                        && At(session.ParaListStrings(doc.Path), 0).Trim().Length == 0
                        ? null
                        : doc.Snapshot();
                }
            });
        }

        private static bool ReadyNumbered(WdSession session, FormatProbe doc, int from, int to)
        {
            return doc.Ready && session.ApplyNumberDefault(doc.Path, from, to) && doc.ReadParagraphs();
        }

        private static string Inner(FormatProbe doc, string needle)
        {
            foreach (Marker marker in doc.Paragraphs)
            {
                if ((marker.Text ?? "").IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return marker.Text ?? "";
                }
            }

            return "";
        }

        private static string ParaInner(FormatProbe doc, string needle)
        {
            foreach (Marker marker in doc.Paragraphs)
            {
                if ((marker.Text ?? "").IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return marker.Text ?? "";
                }
            }

            return "";
        }

        private static bool TagThen(string inner, string body, string listString)
        {
            string text = Regex.Replace(inner ?? "", @"^\[in_table=[^\]]+\]", "");
            Match match = AutonumTag.Match(text);
            if (!match.Success)
            {
                return false;
            }

            string v = Unescape(match.Groups[1].Value).Trim();
            string list = (listString ?? "").Trim();
            if (list.Length > 0 && v != list)
            {
                return false;
            }

            string rest = text.Substring(match.Index + match.Length).TrimEnd('\r', '\n', '\a');
            return rest == body;
        }

        private static string Unescape(string value)
        {
            return (value ?? "")
                .Replace("&quot;", "\"")
                .Replace("&lt;", "<")
                .Replace("&gt;", ">")
                .Replace("&amp;", "&");
        }

        private static string Story(string text)
        {
            return (text ?? "").TrimEnd('\r', '\a', '\n', '\u0007');
        }

        private static string At(List<string> list, int index)
        {
            return list != null && index >= 0 && index < list.Count ? list[index] ?? "" : "";
        }

        private static bool NoTag(string inner)
        {
            return (inner ?? "").IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) < 0;
        }

        private static bool TaggedBody(FormatProbe doc, string body, string list)
        {
            return TagThen(Inner(doc, body), body, list);
        }

        private static string ListFor(List<string> stories, List<string> lists, string body)
        {
            if (stories == null)
            {
                return "";
            }

            for (int i = 0; i < stories.Count; i++)
            {
                if (Story(stories[i]) == body)
                {
                    return At(lists, i);
                }
            }

            return "";
        }

        private static bool StoryHas(List<string> stories, string body)
        {
            if (stories == null)
            {
                return false;
            }

            foreach (string story in stories)
            {
                if (Story(story) == body || (story ?? "").IndexOf(body, StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool NumberBody(WdSession session, string path, string body)
        {
            List<string> stories = session.ParaStories(path);
            for (int i = 0; i < stories.Count; i++)
            {
                if (Story(stories[i]) == body)
                {
                    return session.ApplyNumberDefault(path, i + 1, i + 1);
                }
            }

            return false;
        }

        private static ToolResult ClearCodes(WdSession session, FormatProbe doc, string codes)
        {
            return session.ApplyParagraphFormat(new Dictionary<string, object>
            {
                ["channel_id"] = doc.Id,
                ["target_paragraph_codes"] = codes,
                ["para_format"] = new Dictionary<string, object>
                {
                    ["list"] = new Dictionary<string, object> { ["kind"] = "none" }
                }
            });
        }

        private static List<Dictionary<string, object>> InsertStart(string insert)
        {
            return One("insert", new Dictionary<string, object>
            {
                ["target"] = new Dictionary<string, object> { ["position"] = "start" },
                ["insert"] = insert,
                ["format"] = DefaultFormat()
            });
        }

        private static List<Dictionary<string, object>> Replace(string codes, string neu)
        {
            return One("replace", new Dictionary<string, object>
            {
                ["original"] = new Dictionary<string, object> { ["paragraph_codes"] = codes },
                ["new"] = neu,
                ["format"] = DefaultFormat()
            });
        }

        private static List<Dictionary<string, object>> InsertAfter(string codes, string insert)
        {
            return One("insert", new Dictionary<string, object>
            {
                ["target"] = new Dictionary<string, object> { ["paragraph_codes"] = codes },
                ["insert"] = insert,
                ["format"] = DefaultFormat()
            });
        }

        private static List<Dictionary<string, object>> One(string type, Dictionary<string, object> detail)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = type,
                    ["detail"] = detail
                }
            };
        }

        private static Dictionary<string, object> DefaultFormat()
        {
            return new Dictionary<string, object>
            {
                ["char_format"] = new Dictionary<string, object>
                {
                    ["font_name"] = "宋体",
                    ["font_size"] = 12,
                    ["is_bold"] = false
                }
            };
        }

        private static FormatProbe Open(WdSession session, string text)
        {
            return FormatProbe.Open(session, text, "");
        }

        private static void Case(TestRun run, string id, Func<string> body)
        {
            if (Only != null && !Only.Contains(id))
            {
                return;
            }

            try
            {
                string error = body();
                if (error == null)
                {
                    run.CaseOk(id);
                }
                else
                {
                    run.CaseFail(id, error);
                }
            }
            catch (Exception ex)
            {
                run.CaseFail(id, ex.Message);
            }
        }

        private static HashSet<string> Pick(string filter, string prefix)
        {
            if (string.IsNullOrWhiteSpace(filter)
                || filter.Equals(prefix, StringComparison.OrdinalIgnoreCase)
                || filter.Equals(prefix + "-", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var picked = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string part in filter.Split(','))
            {
                string id = part.Trim();
                if (id.Length > 0)
                {
                    picked.Add(id);
                }
            }

            return picked.Count > 0 ? picked : null;
        }
    }
}
