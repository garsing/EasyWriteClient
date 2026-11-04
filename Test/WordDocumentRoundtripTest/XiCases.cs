using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 抽图副本改到工作区根。前缀 xi-。正式 25 + 巩固 25。环境条不在这里。
    /// </summary>
    internal static class XiCases
    {
        public const int Target = 50;

        private static readonly string[] ConsolidateIds =
        {
            "xi-c-lv-s", "xi-c-lv-s2", "xi-c-lv-p", "xi-c-lv-p2", "xi-c-lv-empty", "xi-c-lv-upper",
            "xi-c-two-both", "xi-c-two-abs", "xi-c-two-no3", "xi-c-two-png",
            "xi-c-pk", "xi-c-not-open", "xi-c-hash", "xi-c-mem-text", "xi-c-disk-text", "xi-c-copy-text",
            "xi-c-fields", "xi-c-chan", "xi-c-verbose", "xi-c-slash",
            "xi-c-third", "xi-c-keep-img", "xi-c-get-mid",
            "xi-c-lv-bad", "xi-c-junk"
        };

        public static void Run(TestRun run, string host)
        {
            var s = new WdSession(host);
            RunSchema(run);
            string root;
            string workspaceError;
            if (!TryWorkspace(out root, out workspaceError))
            {
                Fail(run, "xi-ok-msg", workspaceError);
                Fail(run, "xi-ok-src", workspaceError);
                Fail(run, "xi-ok-files", workspaceError);
                Fail(run, "xi-ok-fullname", workspaceError);
                Fail(run, "xi-ok-dirty", workspaceError);
                Fail(run, "xi-ok-session", workspaceError);
                Fail(run, "xi-ok-tmp", workspaceError);
                Fail(run, "xi-ok-para", workspaceError);
                Fail(run, "xi-ok-sent", workspaceError);
                Fail(run, "xi-ok-host", workspaceError);
                Fail(run, "xi-ov-once", workspaceError);
                Fail(run, "xi-ov-clear", workspaceError);
                Fail(run, "xi-ov-src", workspaceError);
                Fail(run, "xi-ov-other", workspaceError);
                Fail(run, "xi-lite-src", workspaceError);
                Fail(run, "xi-lite-nofile", workspaceError);
                Fail(run, "xi-none-ok", workspaceError);
                Fail(run, "xi-none-files", workspaceError);
                Fail(run, "xi-fail-channel", workspaceError);
                Fail(run, "xi-fail-level", workspaceError);
                Fail(run, "xi-cache-hit", workspaceError);
                Fail(run, "xi-cache-keep", workspaceError);
                Fail(run, "xi-cache-skip", workspaceError);
                Fail(run, "xi-ckpt-dir", workspaceError);
                FailIds(run, ConsolidateIds, workspaceError);
                return;
            }

            string curr = Path.Combine(root.Replace('/', '\\'), "curr_docs");
            string sessions = Path.Combine(root.Replace('/', '\\'), "sessions");
            RunFailChannel(run);
            RunMain(run, s, curr, sessions);
            RunNone(run, s, curr);
            RunFailLevel(run, s, curr);
            RunConsolidate(run, s, root, curr);
        }

        private static void RunSchema(TestRun run)
        {
            string path = FindRepoFile(@"WordAddInBackend\main-server\project\McpTools\F_get_document_content_tool.py");
            if (path == null)
            {
                Fail(run, "xi-lite-schema", "找不到 F_get_document_content_tool.py");
                return;
            }

            string text = File.ReadAllText(path);
            int paramsAt = text.IndexOf("\"parameters\"", StringComparison.Ordinal);
            string parameters = paramsAt >= 0 ? text.Substring(paramsAt) : text;
            bool hasSwitch = parameters.IndexOf("extract", StringComparison.OrdinalIgnoreCase) >= 0;
            Ok(run, "xi-lite-schema", !hasSwitch, "F_get_document_content 参数表里出现了 extract");
        }

        private static void RunFailChannel(TestRun run)
        {
            ToolResult wr = DesktopToolPipe.Call(
                "F_extract_document_images",
                new Dictionary<string, object> { ["channel_id"] = "w999" });
            Ok(run, "xi-fail-channel",
                FailedClean(wr),
                "未知渠道应失败且没有抽取成功。实际 " + Dump(wr));
        }

        private static void RunMain(TestRun run, WdSession s, string curr, string sessions)
        {
            string png = WritePng();
            string path = s.NewDocxPath();
            string otherPath = null;
            try
            {
                ToolResult opened = s.Open(path, true, s.App);
                if (!opened.Success)
                {
                    FailMain(run, "开稿失败: " + (opened.Error ?? ""));
                    return;
                }

                string id = WdSession.Str(opened, "channel_id");
                string stem = Path.GetFileNameWithoutExtension(path);
                ToolResult ins = s.InsertImage(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = png
                });
                if (!ins.Success)
                {
                    FailMain(run, "插图失败: " + (ins.Error ?? ""));
                    return;
                }

                ToolResult lite = s.GetContent(id);
                string liteText = WdSession.Str(lite, "document_content");
                string liteSrc = FirstSrc(liteText);
                Ok(run, "xi-lite-src",
                    lite.Success
                        && liteText.IndexOf("抽取成功", StringComparison.Ordinal) < 0
                        && liteSrc.IndexOf("lightweight/placeholder_", StringComparison.Ordinal) >= 0
                        && liteSrc.IndexOf(":/", StringComparison.Ordinal) < 0,
                    lite.Success ? "简单版 src=" + liteSrc : (lite.Error ?? "读文失败"));
                Ok(run, "xi-lite-nofile",
                    !File.Exists(Path.Combine(curr, stem + ".docx"))
                        && !File.Exists(Path.Combine(curr, stem + "_copy.docx"))
                        && !Directory.Exists(WordImageDir(curr, stem)),
                    "简单版读文后 curr_docs 出现了 " + stem);

                string fullBefore;
                bool savedBefore;
                bool sawState = s.TryReadOpenState(path, out fullBefore, out savedBefore);
                DateTime mtimeBefore = File.Exists(path) ? File.GetLastWriteTimeUtc(path) : DateTime.MinValue;
                string[] ckptBefore = SnapshotCheckpoints();

                ToolResult ext = Extract(id, null);
                string display = WdSession.Str(ext, "document_content");
                string src = FirstSrc(display);
                string fullAfter;
                bool savedAfter;
                bool sawAfter = s.TryReadOpenState(path, out fullAfter, out savedAfter);
                DateTime mtimeAfter = File.Exists(path) ? File.GetLastWriteTimeUtc(path) : DateTime.MinValue;

                Ok(run, "xi-ok-msg",
                    ext.Success
                        && WdSession.Str(ext, "message") == "抽取成功"
                        && !string.IsNullOrEmpty(display)
                        && !HasImageList(ext),
                    ext.Success ? "回包缺抽取成功或带了图片清单。message=" + WdSession.Str(ext, "message") : (ext.Error ?? ""));
                string expectedPrefix = (WordImageDir(curr, stem) + "\\image_1.").Replace('\\', '/');
                bool srcOk = ext.Success
                    && src.IndexOf(":/", StringComparison.Ordinal) >= 0
                    && src.IndexOf('\\') < 0
                    && src.StartsWith(expectedPrefix, StringComparison.OrdinalIgnoreCase)
                    && File.Exists(src.Replace('/', '\\'));
                Ok(run, "xi-ok-src", srcOk, ext.Success ? "src=" + src + " 期望前缀 " + expectedPrefix : (ext.Error ?? ""));
                Ok(run, "xi-ok-files",
                    ext.Success && PairExists(curr, stem) && !HashDir(curr, stem),
                    "副本或 image 目录不在 " + WordImageDir(curr, stem));
                Ok(run, "xi-ok-fullname",
                    sawState && sawAfter && SamePath(fullBefore, fullAfter),
                    "抽前 " + fullBefore + " 抽后 " + fullAfter);
                Ok(run, "xi-ok-dirty",
                    sawState && sawAfter && savedBefore == false && savedAfter == false
                        && mtimeBefore != DateTime.MinValue
                        && mtimeBefore == mtimeAfter,
                    "Saved 抽前=" + savedBefore + " 抽后=" + savedAfter
                        + " 原文件时间 " + mtimeBefore.ToString("o") + " → " + mtimeAfter.ToString("o"));
                Ok(run, "xi-ok-session",
                    !TreeHasStem(sessions, stem),
                    "sessions 里出现了 " + stem);
                Ok(run, "xi-ok-tmp",
                    !TmpHasStem(stem),
                    "插件旁 tmp 里出现了 " + stem);
                Ok(run, "xi-ok-sent",
                    ext.Success && ImageOutsideSentences(display),
                    "句子 display 里没有图，或 <image 进了 S_");
                Ok(run, "xi-ok-host",
                    sawState && sawAfter && SamePath(fullBefore, fullAfter),
                    "活动文档路径变了: " + fullBefore + " → " + fullAfter);

                ToolResult again = s.GetContent(id);
                string againText = WdSession.Str(again, "document_content");
                Ok(run, "xi-cache-hit",
                    again.Success && againText == liteText && againText.IndexOf(":/", StringComparison.Ordinal) < 0,
                    again.Success
                        ? "第二次读文与抽图前不一致，或带了盘符路径"
                        : (again.Error ?? ""));
                Ok(run, "xi-cache-keep",
                    again.Success
                        && FirstSrc(againText).IndexOf("lightweight/placeholder_", StringComparison.Ordinal) >= 0,
                    "第二次读文 src=" + FirstSrc(againText));

                string imageDir = WordImageDir(curr, stem);
                Directory.CreateDirectory(imageDir);
                string leftover = Path.Combine(imageDir, "image_99.png");
                File.WriteAllBytes(leftover, new byte[] { 1, 2, 3 });
                string copy = Path.Combine(curr, stem + "_copy.docx");
                if (File.Exists(copy))
                {
                    File.Delete(copy);
                }

                ToolResult second = Extract(id, "paragraph");
                string para = WdSession.Str(second, "document_content");
                string paraSrc = FirstSrc(para);
                Ok(run, "xi-ok-para",
                    second.Success
                        && WdSession.Str(second, "message") == "抽取成功"
                        && paraSrc.StartsWith(expectedPrefix, StringComparison.OrdinalIgnoreCase)
                        && File.Exists(paraSrc.Replace('/', '\\')),
                    second.Success ? "段 display src=" + paraSrc : (second.Error ?? ""));
                Ok(run, "xi-ov-once",
                    PairExists(curr, stem) && !TimestampExtra(curr, stem),
                    "同名再抽后副本不是唯一一对");
                Ok(run, "xi-ov-clear",
                    !File.Exists(leftover) && HasImage1(imageDir),
                    "image_99 仍在，或没有 image_1");
                Ok(run, "xi-ov-src",
                    second.Success && paraSrc.IndexOf("/image_1.", StringComparison.OrdinalIgnoreCase) >= 0
                        && File.Exists(paraSrc.Replace('/', '\\')),
                    "再抽 src=" + paraSrc);
                Ok(run, "xi-cache-skip",
                    second.Success && File.Exists(copy) && DocumentState.FingerprintVersion == 11,
                    "第二次抽图没有重写 _copy.docx，或 FingerprintVersion=" + DocumentState.FingerprintVersion);

                string[] ckptAfter = SnapshotCheckpoints();
                string ckptDiff = DiffSet(ckptBefore, ckptAfter);
                Ok(run, "xi-ckpt-dir",
                    ckptDiff.Length == 0 && !Directory.Exists(Path.Combine(curr, "document_checkpoints")),
                    "抽图后检查点目录有变化: " + ckptDiff);

                otherPath = s.NewDocxPath();
                ToolResult otherOpen = s.Open(otherPath, true, s.App);
                string otherId = WdSession.Str(otherOpen, "channel_id");
                string otherStem = Path.GetFileNameWithoutExtension(otherPath);
                ToolResult otherIns = otherOpen.Success
                    ? s.InsertImage(new Dictionary<string, object>
                    {
                        ["channel_id"] = otherId,
                        ["path"] = png
                    })
                    : otherOpen;
                ToolResult otherExt = otherIns.Success ? Extract(otherId, null) : otherIns;
                Ok(run, "xi-ov-other",
                    otherExt.Success && PairExists(curr, stem) && PairExists(curr, otherStem),
                    otherExt.Success ? "两套副本没有同时留下" : (otherExt.Error ?? "另一份稿抽图失败"));
            }
            finally
            {
                s.CloseDocByPath(path);
                if (otherPath != null)
                {
                    s.CloseDocByPath(otherPath);
                }

                TryDelete(png);
            }
        }

        private static void RunNone(TestRun run, WdSession s, string curr)
        {
            string path = s.NewDocxPath();
            try
            {
                ToolResult opened = s.Open(path, true, s.App);
                if (!opened.Success)
                {
                    Fail(run, "xi-none-ok", "开空白稿失败: " + (opened.Error ?? ""));
                    Fail(run, "xi-none-files", "开空白稿失败");
                    return;
                }

                string id = WdSession.Str(opened, "channel_id");
                string stem = Path.GetFileNameWithoutExtension(path);
                ToolResult ext = Extract(id, null);
                Ok(run, "xi-none-ok",
                    ext.Success
                        && WdSession.Str(ext, "message") == "抽取成功"
                        && !string.IsNullOrEmpty(WdSession.Str(ext, "document_content")),
                    ext.Error ?? "空白稿没有抽取成功");
                Ok(run, "xi-none-files",
                    PairExists(curr, stem) && !HasImage1(WordImageDir(curr, stem)),
                    "空白稿副本缺失，或写出了 image_N");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunFailLevel(TestRun run, WdSession s, string curr)
        {
            string path = s.NewDocxPath();
            try
            {
                ToolResult opened = s.Open(path, true, s.App);
                if (!opened.Success)
                {
                    Fail(run, "xi-fail-level", "开稿失败: " + (opened.Error ?? ""));
                    return;
                }

                string id = WdSession.Str(opened, "channel_id");
                string stem = Path.GetFileNameWithoutExtension(path);
                ToolResult wr = Extract(id, "chunk");
                Ok(run, "xi-fail-level",
                    FailedClean(wr)
                        && (wr.Error ?? "").IndexOf("code_level", StringComparison.OrdinalIgnoreCase) >= 0
                        && !File.Exists(Path.Combine(curr, stem + ".docx")),
                    "非法 code_level 应失败且不写副本。实际 " + Dump(wr));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunConsolidate(TestRun run, WdSession s, string root, string curr)
        {
            const string token = "XI巩固甲";
            string png = WritePng();
            string path = s.NewDocxPath();
            string otherPath = null;
            string badPath = null;
            try
            {
                ToolResult opened = s.Open(path, true, s.App);
                if (!opened.Success || !s.TypePlain(path, token + "。"))
                {
                    FailIds(run, ConsolidateIds, opened.Success ? "COM 写句子失败" : ("开稿失败: " + (opened.Error ?? "")));
                    return;
                }

                string id = WdSession.Str(opened, "channel_id");
                string stem = Path.GetFileNameWithoutExtension(path);
                ToolResult ins1 = s.InsertImage(new Dictionary<string, object> { ["channel_id"] = id, ["path"] = png });
                ToolResult ins2 = ins1.Success
                    ? s.InsertImage(new Dictionary<string, object> { ["channel_id"] = id, ["path"] = png })
                    : ins1;
                if (!ins2.Success)
                {
                    FailIds(run, ConsolidateIds, "插两张图失败: " + (ins2.Error ?? ins1.Error ?? ""));
                    return;
                }

                DateTime mtimeBefore = File.Exists(path) ? File.GetLastWriteTimeUtc(path) : DateTime.MinValue;
                string fullBefore;
                bool savedBefore;
                s.TryReadOpenState(path, out fullBefore, out savedBefore);
                string rootSlash = root.Replace('\\', '/').TrimEnd('/');

                ToolResult ext = Extract(id, "s");
                string display = WdSession.Str(ext, "document_content");
                List<string> srcs = AllSrcs(display);
                string imageDir = WordImageDir(curr, stem);
                string docx = Path.Combine(curr, stem + ".docx");
                string copy = Path.Combine(curr, stem + "_copy.docx");
                string fullAfter;
                bool savedAfter;
                s.TryReadOpenState(path, out fullAfter, out savedAfter);

                Ok(run, "xi-c-lv-s",
                    ext.Success && WdSession.Str(ext, "message") == "抽取成功" && WdSession.Str(ext, "code_level") == "sentence",
                    Dump(ext));
                Ok(run, "xi-c-two-both",
                    srcs.Count >= 2
                        && File.Exists(srcs[0].Replace('/', '\\'))
                        && File.Exists(srcs[1].Replace('/', '\\')),
                    "src 数=" + srcs.Count);
                Ok(run, "xi-c-two-abs",
                    srcs.Count >= 2
                        && srcs[0].IndexOf("/" + stem + "_word/image_1.", StringComparison.OrdinalIgnoreCase) >= 0
                        && srcs[1].IndexOf("/" + stem + "_word/image_2.", StringComparison.OrdinalIgnoreCase) >= 0
                        && File.Exists(srcs[0].Replace('/', '\\'))
                        && File.Exists(srcs[1].Replace('/', '\\')),
                    string.Join(" , ", srcs.ToArray()));
                Ok(run, "xi-c-two-no3",
                    srcs.Count == 2 && Directory.Exists(imageDir) && Directory.GetFiles(imageDir, "image_3.*").Length == 0,
                    "src 数=" + srcs.Count);
                Ok(run, "xi-c-two-png",
                    srcs.Count >= 2
                        && StartsWith(srcs[0].Replace('/', '\\'), new byte[] { 0x89, 0x50, 0x4E, 0x47 })
                        && StartsWith(srcs[1].Replace('/', '\\'), new byte[] { 0x89, 0x50, 0x4E, 0x47 }),
                    "落盘文件不是 PNG");
                Ok(run, "xi-c-pk",
                    StartsWith(docx, new byte[] { 0x50, 0x4B }) && StartsWith(copy, new byte[] { 0x50, 0x4B }),
                    "副本不是 zip 文档");
                Ok(run, "xi-c-not-open",
                    savedBefore == false && savedAfter == false
                        && !string.IsNullOrEmpty(fullAfter)
                        && fullAfter.IndexOf("curr_docs", StringComparison.OrdinalIgnoreCase) < 0
                        && SamePath(fullBefore, fullAfter),
                    "Saved " + savedBefore + "→" + savedAfter + " " + fullBefore + " → " + fullAfter);
                DateTime mtimeAfter = File.Exists(path) ? File.GetLastWriteTimeUtc(path) : DateTime.MinValue;
                Ok(run, "xi-c-hash",
                    mtimeBefore != DateTime.MinValue && mtimeBefore == mtimeAfter,
                    "原文件时间 " + mtimeBefore.ToString("o") + " → " + mtimeAfter.ToString("o"));
                Ok(run, "xi-c-mem-text",
                    display.IndexOf(token, StringComparison.Ordinal) >= 0,
                    "抽图 display 里没有未保存句子");
                Ok(run, "xi-c-disk-text",
                    !ZipContains(path, token),
                    "未保存句子写进了原文件");
                Ok(run, "xi-c-copy-text",
                    ZipContains(copy, token),
                    "副本里没有内存中的句子");
                Ok(run, "xi-c-fields",
                    WdSession.Str(ext, "processing_method").Length > 0
                        && string.Equals(WdSession.Str(ext, "host"), s.Host, StringComparison.OrdinalIgnoreCase),
                    "host=" + WdSession.Str(ext, "host") + " method=" + WdSession.Str(ext, "processing_method"));
                Ok(run, "xi-c-chan",
                    WdSession.Str(ext, "channel_id") == id && WdSession.IsShortId(id),
                    "回包 channel_id=" + WdSession.Str(ext, "channel_id") + " 期望 " + id);
                Ok(run, "xi-c-slash",
                    srcs.Count >= 2
                        && srcs[0].StartsWith(rootSlash + "/curr_docs/", StringComparison.OrdinalIgnoreCase)
                        && srcs[0].IndexOf('\\') < 0
                        && srcs[1].IndexOf('\\') < 0,
                    string.Join(" , ", srcs.ToArray()));

                ToolResult mid = s.GetContent(id);
                string midSrc = FirstSrc(WdSession.Str(mid, "document_content"));
                Ok(run, "xi-c-get-mid",
                    mid.Success && midSrc.IndexOf("lightweight/placeholder_", StringComparison.Ordinal) >= 0,
                    "中途读文 src=" + midSrc);

                ToolResult upper = Extract(id, "SENTENCE", true);
                string upperText = WdSession.Str(upper, "document_content");
                Ok(run, "xi-c-lv-upper",
                    upper.Success && WdSession.Str(upper, "code_level") == "sentence" && WdSession.Str(upper, "message") == "抽取成功",
                    Dump(upper));
                Ok(run, "xi-c-verbose",
                    upper.Success && AllSrcs(upperText).Count == 2 && upperText.IndexOf(token, StringComparison.Ordinal) >= 0,
                    "verbose 抽图 src 数=" + AllSrcs(upperText).Count);

                Directory.CreateDirectory(imageDir);
                string junkTxt = Path.Combine(imageDir, "note.txt");
                string junkImg = Path.Combine(imageDir, "image_9.png");
                File.WriteAllText(junkTxt, "x");
                File.WriteAllBytes(junkImg, new byte[] { 9 });
                ToolResult again = Extract(id, "s_");
                Ok(run, "xi-c-lv-s2",
                    again.Success && WdSession.Str(again, "code_level") == "sentence",
                    Dump(again));
                Ok(run, "xi-c-junk",
                    !File.Exists(junkTxt) && !File.Exists(junkImg)
                        && Directory.Exists(imageDir) && Directory.GetFiles(imageDir, "image_1.*").Length == 1,
                    "残留 note.txt 或 image_9 还在");
                Ok(run, "xi-c-third",
                    PairExists(curr, stem) && !TimestampExtra(curr, stem),
                    "再抽后不是唯一一对副本");

                ToolResult emptyLv = Extract(id, "");
                Ok(run, "xi-c-lv-empty",
                    emptyLv.Success && WdSession.Str(emptyLv, "code_level") == "sentence" && WdSession.Str(emptyLv, "message") == "抽取成功",
                    Dump(emptyLv));
                ToolResult para = Extract(id, "p");
                Ok(run, "xi-c-lv-p",
                    para.Success
                        && WdSession.Str(para, "code_level") == "paragraph"
                        && FirstSrc(WdSession.Str(para, "document_content")).IndexOf("/image_1.", StringComparison.OrdinalIgnoreCase) >= 0,
                    Dump(para));
                ToolResult para2 = Extract(id, "p_");
                Ok(run, "xi-c-lv-p2",
                    para2.Success && WdSession.Str(para2, "code_level") == "paragraph" && WdSession.Str(para2, "message") == "抽取成功",
                    Dump(para2));

                badPath = s.NewDocxPath();
                ToolResult badOpen = s.Open(badPath, true, s.App);
                string badId = WdSession.Str(badOpen, "channel_id");
                string badStem = Path.GetFileNameWithoutExtension(badPath);
                ToolResult bad = badOpen.Success ? Extract(badId, "1") : badOpen;
                Ok(run, "xi-c-lv-bad",
                    badOpen.Success
                        && FailedClean(bad)
                        && (bad.Error ?? "").IndexOf("code_level", StringComparison.OrdinalIgnoreCase) >= 0
                        && !File.Exists(Path.Combine(curr, badStem + ".docx")),
                    Dump(bad));

                otherPath = s.NewDocxPath();
                ToolResult otherOpen = s.Open(otherPath, true, s.App);
                string otherId = WdSession.Str(otherOpen, "channel_id");
                ToolResult otherIns = otherOpen.Success
                    ? s.InsertImage(new Dictionary<string, object> { ["channel_id"] = otherId, ["path"] = png })
                    : otherOpen;
                ToolResult otherExt = otherIns.Success ? Extract(otherId, null) : otherIns;
                Ok(run, "xi-c-keep-img",
                    otherExt.Success
                        && Directory.Exists(imageDir)
                        && Directory.GetFiles(imageDir, "image_1.*").Length == 1
                        && Directory.GetFiles(imageDir, "image_2.*").Length == 1,
                    otherExt.Success ? "另一份稿抽完后，前一份的 image_1/image_2 丢了" : (otherExt.Error ?? ""));
            }
            finally
            {
                s.CloseDocByPath(path);
                if (otherPath != null)
                {
                    s.CloseDocByPath(otherPath);
                }

                if (badPath != null)
                {
                    s.CloseDocByPath(badPath);
                }

                TryDelete(png);
            }
        }

        private static void FailMain(TestRun run, string reason)
        {
            string[] ids =
            {
                "xi-ok-msg", "xi-ok-src", "xi-ok-files", "xi-ok-fullname", "xi-ok-dirty",
                "xi-ok-session", "xi-ok-tmp", "xi-ok-para", "xi-ok-sent", "xi-ok-host",
                "xi-ov-once", "xi-ov-clear", "xi-ov-src", "xi-ov-other",
                "xi-lite-src", "xi-lite-nofile",
                "xi-cache-hit", "xi-cache-keep", "xi-cache-skip", "xi-ckpt-dir"
            };
            foreach (string id in ids)
            {
                Fail(run, id, reason);
            }
        }

        private static ToolResult Extract(string channelId, string codeLevel)
        {
            return Extract(channelId, codeLevel, false);
        }

        private static ToolResult Extract(string channelId, string codeLevel, bool verbose)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            if (codeLevel != null)
            {
                args["code_level"] = codeLevel;
            }

            if (verbose)
            {
                args["verbose"] = true;
            }

            return DesktopToolPipe.Call("F_extract_document_images", args);
        }

        private static bool TryWorkspace(out string root, out string error)
        {
            root = null;
            error = null;
            ToolResult info = DesktopToolPipe.Call("F_workspace_info", new Dictionary<string, object>());
            root = WdSession.Str(info, "workspace_root");
            if (!info.Success || string.IsNullOrWhiteSpace(root))
            {
                error = info.Error ?? "工作区根不可用";
                return false;
            }

            return true;
        }

        private static bool FailedClean(ToolResult wr)
        {
            if (wr == null || wr.Success)
            {
                return false;
            }

            if (!string.IsNullOrEmpty(WdSession.Str(wr, "document_content")))
            {
                return false;
            }

            string message = WdSession.Str(wr, "message");
            return message.IndexOf("抽取成功", StringComparison.Ordinal) < 0
                && (wr.Error ?? "").IndexOf("抽取成功", StringComparison.Ordinal) < 0;
        }

        private static bool HasImageList(ToolResult wr)
        {
            Dictionary<string, object> data = WdSession.Data(wr);
            if (data == null)
            {
                return false;
            }

            return data.ContainsKey("images")
                || data.ContainsKey("image_files")
                || data.ContainsKey("image_list")
                || data.ContainsKey("files");
        }

        private static List<string> AllSrcs(string text)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(text))
            {
                return list;
            }

            foreach (Match m in Regex.Matches(text, "src=\"([^\"]*)\""))
            {
                list.Add(m.Groups[1].Value);
            }

            return list;
        }

        private static bool StartsWith(string path, byte[] magic)
        {
            try
            {
                if (!File.Exists(path) || magic == null)
                {
                    return false;
                }

                byte[] head = new byte[magic.Length];
                using (FileStream fs = File.OpenRead(path))
                {
                    if (fs.Read(head, 0, head.Length) != head.Length)
                    {
                        return false;
                    }
                }

                for (int i = 0; i < magic.Length; i++)
                {
                    if (head[i] != magic[i])
                    {
                        return false;
                    }
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool ZipContains(string path, string token)
        {
            try
            {
                using (var zip = System.IO.Compression.ZipFile.OpenRead(path))
                {
                    System.IO.Compression.ZipArchiveEntry entry = zip.GetEntry("word/document.xml");
                    if (entry == null)
                    {
                        return false;
                    }

                    using (Stream stream = entry.Open())
                    using (var reader = new StreamReader(stream))
                    {
                        string xml = reader.ReadToEnd();
                        return xml.IndexOf(token, StringComparison.Ordinal) >= 0;
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string FirstSrc(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            Match m = Regex.Match(text, "src=\"([^\"]*)\"");
            return m.Success ? m.Groups[1].Value : "";
        }

        private static bool ImageOutsideSentences(string text)
        {
            if (string.IsNullOrEmpty(text) || text.IndexOf("<image", StringComparison.Ordinal) < 0)
            {
                return false;
            }

            foreach (Match m in Regex.Matches(
                text,
                @"\[S_[^\]]+-start\](.*?)\[S_[^\]]+-end\]",
                RegexOptions.Singleline))
            {
                if (m.Groups[1].Value.IndexOf("<image", StringComparison.Ordinal) >= 0)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool PairExists(string curr, string stem)
        {
            return File.Exists(Path.Combine(curr, stem + ".docx"))
                && File.Exists(Path.Combine(curr, stem + "_copy.docx"))
                && new FileInfo(Path.Combine(curr, stem + ".docx")).Length > 0
                && new FileInfo(Path.Combine(curr, stem + "_copy.docx")).Length > 0;
        }

        private static string WordImageDir(string curr, string stem)
        {
            return Path.Combine(curr, stem + "_word");
        }

        private static bool HashDir(string curr, string stem)
        {
            string dir = WordImageDir(curr, stem);
            if (!Directory.Exists(dir))
            {
                return true;
            }

            string[] files = Directory.GetFiles(dir, "image_1.*");
            return files.Length == 0;
        }

        private static bool HasImage1(string dir)
        {
            return Directory.Exists(dir) && Directory.GetFiles(dir, "image_1.*").Length > 0;
        }

        private static bool TimestampExtra(string curr, string stem)
        {
            if (!Directory.Exists(curr))
            {
                return true;
            }

            foreach (string file in Directory.GetFiles(curr))
            {
                string name = Path.GetFileName(file);
                if (name.IndexOf(stem, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                if (!string.Equals(name, stem + ".docx", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(name, stem + "_copy.docx", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TreeHasStem(string dir, string stem)
        {
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
            {
                return false;
            }

            foreach (string entry in Directory.EnumerateFileSystemEntries(dir, "*" + stem + "*", SearchOption.AllDirectories))
            {
                if (entry.IndexOf(stem, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TmpHasStem(string stem)
        {
            string[] roots =
            {
                @"D:\project\EasyWrite\EasyWriteClient\Desktop\tmp",
                @"D:\project\EasyWrite\EasyWriteClient\Desktop\bin\Debug\net472\tmp",
                @"D:\project\EasyWrite\EasyWriteClient\Desktop\bin\Release\net472\tmp",
                @"D:\project\EasyWrite\EasyWriteClient\Plugin\tmp",
                @"D:\project\EasyWrite\EasyWriteClient\Plugin\bin\Debug\net472\tmp"
            };
            foreach (string root in roots)
            {
                if (TreeHasStem(root, stem))
                {
                    return true;
                }
            }

            return false;
        }

        private static string[] SnapshotCheckpoints()
        {
            var list = new List<string>();
            string[] roots =
            {
                @"D:\project\EasyWrite\EasyWriteClient\Desktop\bin\tmp",
                @"D:\project\EasyWrite\EasyWriteClient\Desktop\tmp",
                @"D:\project\EasyWrite\EasyWriteClient\Desktop\bin\Debug\net472\tmp",
                @"D:\project\EasyWrite\EasyWriteClient\Plugin\tmp",
                @"D:\project\EasyWrite\EasyWriteClient\Plugin\bin\Debug\net472\tmp"
            };
            foreach (string root in roots)
            {
                if (!Directory.Exists(root))
                {
                    continue;
                }

                foreach (string dir in Directory.EnumerateDirectories(root, "document_checkpoints", SearchOption.AllDirectories))
                {
                    foreach (string file in Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories))
                    {
                        list.Add(file);
                    }
                }
            }

            list.Sort(StringComparer.OrdinalIgnoreCase);
            return list.ToArray();
        }

        private static string DiffSet(string[] before, string[] after)
        {
            var oldSet = new HashSet<string>(before, StringComparer.OrdinalIgnoreCase);
            var newSet = new HashSet<string>(after, StringComparer.OrdinalIgnoreCase);
            var lines = new List<string>();
            foreach (string path in newSet)
            {
                if (!oldSet.Contains(path))
                {
                    lines.Add("新增 " + path);
                }
            }

            foreach (string path in oldSet)
            {
                if (!newSet.Contains(path))
                {
                    lines.Add("消失 " + path);
                }
            }

            if (lines.Count == 0)
            {
                return "";
            }

            int n = Math.Min(lines.Count, 4);
            return string.Join(" | ", lines.GetRange(0, n).ToArray());
        }

        private static bool SamePath(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
            {
                return false;
            }

            return string.Equals(
                Path.GetFullPath(a).TrimEnd('\\'),
                Path.GetFullPath(b).TrimEnd('\\'),
                StringComparison.OrdinalIgnoreCase);
        }

        private static string WritePng()
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-xi-" + Guid.NewGuid().ToString("N") + ".png");
            using (var bmp = new System.Drawing.Bitmap(32, 16))
            using (var g = System.Drawing.Graphics.FromImage(bmp))
            {
                g.Clear(System.Drawing.Color.FromArgb(0x44, 0x72, 0xC4));
                bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
            }

            return path;
        }

        private static string FindRepoFile(string relative)
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            for (int i = 0; i < 8 && !string.IsNullOrEmpty(dir); i++)
            {
                string candidate = Path.Combine(dir, relative);
                if (File.Exists(candidate))
                {
                    return candidate;
                }

                DirectoryInfo parent = Directory.GetParent(dir);
                dir = parent == null ? null : parent.FullName;
            }

            string fallback = Path.Combine(@"D:\project\EasyWrite", relative);
            return File.Exists(fallback) ? fallback : null;
        }

        private static string Dump(ToolResult wr)
        {
            if (wr == null)
            {
                return "(null)";
            }

            return "success=" + wr.Success + " error=" + (wr.Error ?? "") + " message=" + WdSession.Str(wr, "message");
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        private static void Ok(TestRun run, string id, bool cond, string reason)
        {
            if (cond)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, reason);
            }
        }

        private static void Fail(TestRun run, string id, string reason)
        {
            run.CaseFail(id, reason);
        }

        private static void FailIds(TestRun run, string[] ids, string reason)
        {
            foreach (string id in ids)
            {
                run.CaseFail(id, reason);
            }
        }
    }
}
