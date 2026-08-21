using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using WordAddIn1.Terminal;

namespace WordAddIn1
{
    public static class F_RunTerminalTool
    {
        private static readonly Regex NeedsPython = new Regex(
            @"(?i)(\bpython\b|\bpy\.exe\b|\.py\b|python-docx)",
            RegexOptions.Compiled);

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_run_terminal"] = RunAsync;
        }

        private static async Task<ToolResult> RunAsync(Dictionary<string, object> args)
        {
            try
            {
                AgentRunCancellation.ThrowIfCancelled();
                if (ChannelHost.Kind != ChannelHostKind.Desktop)
                {
                    return Fail("仅易写桌面应用支持");
                }

                string command = OpenDocumentPath.TryGetString(args, "command");
                if (string.IsNullOrWhiteSpace(command))
                {
                    return Fail("必须提供 command 参数");
                }

                string cwdArg = OpenDocumentPath.TryGetString(args, "cwd");
                if (OpenDocumentCommandGuard.TryFindBlockedPath(command, cwdArg, out string blocked))
                {
                    return Fail("文件仍在办公软件中打开：" + blocked + "。请先 F_close_document，再执行终端命令。");
                }

                int timeoutSec = ParseTimeout(args);
                bool useManaged = OpenDocumentPath.ParseBool(args, "use_managed_python", false)
                    || NeedsPython.IsMatch(command);

                string pythonExe = null;
                if (useManaged)
                {
                    try
                    {
                        await ManagedPythonRuntime.EnsureInstalledAsync(AgentRunCancellation.Token)
                            .ConfigureAwait(false);
                        pythonExe = ManagedPythonRuntime.PythonExe;
                    }
                    catch (OperationCanceledException)
                    {
                        return Fail("cancelled by user");
                    }
                    catch (Exception ex)
                    {
                        return Fail("托管 Python 安装失败: " + ex.Message);
                    }
                }

                string conversationId = ConversationContext.CurrentId;
                string defaultCwd = WorkspacePathResolver.GetSessionDirectory(conversationId);
                string openCwd = string.IsNullOrWhiteSpace(cwdArg) ? defaultCwd : cwdArg;

                TerminalSession session = TerminalSessionHub.GetOrCreate(conversationId, openCwd);
                if (useManaged && !session.ManagedPythonPathPrepended)
                {
                    string pathPrefix = System.IO.Path.GetDirectoryName(ManagedPythonRuntime.PythonExe);
                    session.RunBootstrap(
                        "$env:PATH = '" + pathPrefix.Replace("'", "''") + ";' + $env:PATH; "
                        + "$env:PYTHONPATH = '" + ManagedPythonRuntime.SiteDirectory.Replace("'", "''") + "'");
                    session.ManagedPythonPathPrepended = true;
                }

                using (var linked = CancellationTokenSource.CreateLinkedTokenSource(AgentRunCancellation.Token))
                {
                    linked.Token.Register(() => session.KillCurrentCommand());
                    TerminalCommandResult result = await session.RunCommandAsync(
                        command, cwdArg, timeoutSec, linked.Token).ConfigureAwait(false);

                    var data = new Dictionary<string, object>
                    {
                        ["exit_code"] = result.ExitCode,
                        ["stdout"] = result.Stdout ?? "",
                        ["stderr"] = result.Stderr ?? "",
                        ["truncated"] = result.Truncated,
                        ["cwd"] = result.Cwd ?? session.Cwd
                    };
                    if (!string.IsNullOrEmpty(pythonExe))
                    {
                        data["python_exe"] = pythonExe;
                    }

                    if (result.Cancelled)
                    {
                        return new ToolResult { Success = false, Error = "cancelled by user", Data = data };
                    }

                    if (result.TimedOut)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "命令超时（" + timeoutSec + " 秒）",
                            Data = data
                        };
                    }

                    bool ok = result.ExitCode == 0;
                    return new ToolResult
                    {
                        Success = ok,
                        Error = ok ? null : "exit_code=" + result.ExitCode,
                        Data = data
                    };
                }
            }
            catch (OperationCanceledException)
            {
                return Fail("cancelled by user");
            }
            catch (Exception ex)
            {
                return Fail("终端执行失败: " + ex.Message);
            }
        }

        private static int ParseTimeout(Dictionary<string, object> args)
        {
            int value = 120;
            if (args != null && args.ContainsKey("timeout_sec") && args["timeout_sec"] != null)
            {
                int.TryParse(args["timeout_sec"].ToString(), out value);
            }

            if (value <= 0)
            {
                value = 120;
            }

            return Math.Min(value, 300);
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }
    }
}
