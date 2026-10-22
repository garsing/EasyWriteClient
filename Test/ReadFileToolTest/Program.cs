using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace ReadFileToolTest
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            string filter = GetArg(args, "--case=");
            Console.WriteLine("F_read_file 短套（rf-*），直调 Shared 注册表");
            Console.WriteLine();

            var registry = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(StringComparer.OrdinalIgnoreCase);
            F_ReadFileTool.Register(registry, null);
            var run = new TestRun();
            var sw = Stopwatch.StartNew();
            using (var fx = new Fixtures())
            {
                Cases.Run(run, registry, fx, filter);
            }

            sw.Stop();
            Console.WriteLine();
            Console.WriteLine(
                "过 " + run.CasesPassed + " / 败 " + run.CasesFailed + " / 跳 " + run.CasesSkipped
                + "  用时 " + sw.ElapsedMilliseconds + " ms");
            return run.CasesFailed > 0 ? 1 : 0;
        }

        internal static ToolResult Call(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> registry,
            Dictionary<string, object> args)
        {
            return registry["F_read_file"](args).GetAwaiter().GetResult();
        }

        internal static JObject Data(ToolResult r)
        {
            if (r == null || r.Data == null)
            {
                return new JObject();
            }

            if (r.Data is JObject jo)
            {
                return jo;
            }

            if (r.Data is Dictionary<string, object> d)
            {
                return JObject.FromObject(d);
            }

            return JObject.FromObject(r.Data);
        }

        internal static string GetArg(string[] args, string prefix)
        {
            if (args == null)
            {
                return "";
            }

            foreach (string a in args)
            {
                if (a != null && a.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    return a.Substring(prefix.Length);
                }
            }

            return "";
        }
    }
}
