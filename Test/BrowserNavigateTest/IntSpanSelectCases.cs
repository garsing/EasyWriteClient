using System;
using WordAddIn1;

namespace BrowserNavigateTest
{
    /// <summary>Layui SPAN 假下拉的 select。短套，不重跑已绿 int-* 160。</summary>
    internal static class IntSpanSelectCases
    {
        public static void Run(TestRun run, FixtureHttp http, string navHost)
        {
            IntCases.NavHost = navHost;
            var s = new NavSession();
            string id = IntCases.OpenOk(run, s, http.Url("layui-select", "index.html"), "int-span-select-text");
            if (id == null)
            {
                run.CaseFail("int-span-select-contain", "依赖 layui-select");
                run.CaseFail("int-span-select-missing", "依赖 layui-select");
                return;
            }

            string t = IntCases.WaitTree(s, id, true, "combobox", "请选择", "学历");
            string extra = IntCases.Tree(s.Snap(id, domSupplement: "input|select"));
            if (!string.IsNullOrEmpty(extra))
            {
                t = extra;
            }

            string edu = IntCases.FindRef(t, "请选择", "combobox")
                ?? IntCases.FindRef(t, "combobox")
                ?? IntCases.FindRef(t, "请选择")
                ?? IntCases.FindRef(t, "硕士");
            ToolResult r1 = string.IsNullOrEmpty(edu) ? null : s.Interact(id, "select", edu, option: "硕士");
            string t1 = IntCases.Tree(r1);
            IntCases.Ok(
                run,
                "int-span-select-text",
                r1 != null && r1.Success && (t1.IndexOf("硕士", StringComparison.Ordinal) >= 0),
                r1 == null ? "无学历 combobox" : (r1.Success ? t1 : r1.Error));

            string t2 = r1 != null && r1.Success && t1.Length > 0
                ? t1
                : IntCases.Tree(s.Snap(id, domSupplement: "input|select"));
            string rel = LastCombobox(t2) ?? IntCases.FindRef(t2, "股权关系") ?? IntCases.FindRef(t2, "请选择");
            ToolResult r2 = string.IsNullOrEmpty(rel) ? null : s.Interact(id, "select", rel, option: "股权");
            string got2 = IntCases.Tree(r2);
            IntCases.Ok(
                run,
                "int-span-select-contain",
                r2 != null && r2.Success && got2.IndexOf("股权关系", StringComparison.Ordinal) >= 0,
                r2 == null ? "无关系 combobox" : (r2.Success ? got2 : r2.Error));

            string missRef = rel ?? edu;
            IntCases.ExpectFail(
                run,
                "int-span-select-missing",
                string.IsNullOrEmpty(missRef) ? null : s.Interact(id, "select", missRef, option: "不存在的项"),
                "可选");
        }

        private static string LastCombobox(string tree)
        {
            if (string.IsNullOrEmpty(tree))
            {
                return null;
            }

            string last = null;
            foreach (string line in tree.Split('\n'))
            {
                if (line.IndexOf("combobox", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                var m = System.Text.RegularExpressions.Regex.Match(line, @"ref=(e\d+)");
                if (m.Success)
                {
                    last = m.Groups[1].Value;
                }
            }

            return last;
        }
    }
}
