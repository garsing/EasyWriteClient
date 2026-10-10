using System;
using System.Collections.Generic;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// T3：清单工具契约失败应 skip-ckpt-parse，不建副本。
    /// </summary>
    internal static class ContractSkipCkptProbe
    {
        public static int Run()
        {
            Console.WriteLine("契约跳副本探针（ckpt-skip-parse）");
            if (!DesktopToolPipe.TryWaitUntilAvailable(8000))
            {
                Console.WriteLine("FAIL 连不上 Desktop");
                return 1;
            }

            int fail = 0;
            fail += Check(
                "delete_table",
                "F_delete_table",
                new Dictionary<string, object>(),
                "未提供表格编号");
            fail += Check(
                "insert_image",
                "F_insert_image",
                new Dictionary<string, object>(),
                "必须提供 path");
            fail += Check(
                "create_table",
                "F_create_table_from_xml",
                new Dictionary<string, object>(),
                "必须提供 path");
            fail += Check(
                "page_setup",
                "F_apply_page_setup",
                new Dictionary<string, object>
                {
                    ["apply_page_setup"] = false,
                    ["apply_headers_footers"] = false
                },
                "不能同时为 false");
            fail += Check(
                "blank_page",
                "F_insert_blank_page",
                new Dictionary<string, object>
                {
                    ["position"] = "start",
                    ["page_number"] = 2
                },
                "不能同时传");

            Console.WriteLine(fail == 0 ? "契约跳副本: OK" : "契约跳副本: FAIL " + fail);
            return fail == 0 ? 0 : 1;
        }

        private static int Check(
            string name,
            string tool,
            Dictionary<string, object> args,
            string expect)
        {
            ToolResult r = DesktopToolPipe.Call(tool, args);
            bool ok = r != null && !r.Success
                && !string.IsNullOrEmpty(r.Error)
                && r.Error.IndexOf(expect, StringComparison.Ordinal) >= 0;
            Console.WriteLine(
                (ok ? "  [OK] " : "  [FAIL] ") + name + " err=" + (r == null ? "(null)" : r.Error));
            return ok ? 0 : 1;
        }
    }
}
