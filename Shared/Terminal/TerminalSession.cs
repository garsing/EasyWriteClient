using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace WordAddIn1.Terminal
{
    internal sealed class TerminalCommandResult
    {
        public int ExitCode { get; set; }
        public string Stdout { get; set; }
        public string Stderr { get; set; }
        public bool Truncated { get; set; }
        public int StdoutChars { get; set; }
        public int StderrChars { get; set; }
        public string Cwd { get; set; }
        public bool TimedOut { get; set; }
        public bool Cancelled { get; set; }
    }

    internal sealed class TerminalSession : IDisposable
    {
        private const int MaxOutputChars = 32 * 1024;
        private readonly object _gate = new object();
        private Process _process;
        private StreamWriter _stdin;
        private readonly StringBuilder _stdoutBuf = new StringBuilder();
        private readonly StringBuilder _stderrBuf = new StringBuilder();
        private volatile bool _disposed;

        public string Cwd { get; private set; }

        public bool ManagedPythonPathPrepended { get; set; }

        public static TerminalSession Start(string cwd)
        {
            Directory.CreateDirectory(cwd);
            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -NoExit",
                WorkingDirectory = cwd,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };
            psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";

            var session = new TerminalSession { Cwd = cwd };
            session._process = Process.Start(psi);
            if (session._process == null)
            {
                throw new InvalidOperationException("无法启动 PowerShell");
            }

            session._stdin = new StreamWriter(session._process.StandardInput.BaseStream, new UTF8Encoding(false))
            {
                AutoFlush = true
            };
            session._process.OutputDataReceived += (s, e) =>
            {
                if (e.Data != null)
                {
                    lock (session._gate)
                    {
                        session._stdoutBuf.AppendLine(e.Data);
                    }
                }
            };
            session._process.ErrorDataReceived += (s, e) =>
            {
                if (e.Data != null)
                {
                    lock (session._gate)
                    {
                        session._stderrBuf.AppendLine(e.Data);
                    }
                }
            };
            session._process.BeginOutputReadLine();
            session._process.BeginErrorReadLine();

            session._stdin.WriteLine("$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new()");
            return session;
        }

        public void RunBootstrap(string command)
        {
            if (_disposed || _stdin == null)
            {
                return;
            }

            _stdin.WriteLine(command ?? "");
        }

        public async Task<TerminalCommandResult> RunCommandAsync(
            string command,
            string cwd,
            int timeoutSec,
            CancellationToken cancellationToken)
        {
            if (_disposed || _process == null || _process.HasExited)
            {
                throw new InvalidOperationException("终端会话已关闭");
            }

            if (!string.IsNullOrWhiteSpace(cwd)
                && !string.Equals(cwd, Cwd, StringComparison.OrdinalIgnoreCase))
            {
                _stdin.WriteLine("Set-Location -LiteralPath '" + EscapePs(cwd) + "'");
                Cwd = cwd;
            }

            string marker = "EWTERM_" + Guid.NewGuid().ToString("N");
            string encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(command ?? ""));
            lock (_gate)
            {
                _stdoutBuf.Clear();
                _stderrBuf.Clear();
            }

            string outFile = Path.Combine(Path.GetTempPath(), marker + ".out");
            string errFile = Path.Combine(Path.GetTempPath(), marker + ".err");
            _stdin.WriteLine(
                "$__ew_out = '" + EscapePs(outFile) + "'; $__ew_err = '" + EscapePs(errFile) + "'; "
                + "$__ew_p = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-EncodedCommand','"
                + encoded
                + "') -WorkingDirectory (Get-Location).Path -Wait -PassThru -WindowStyle Hidden "
                + "-RedirectStandardOutput $__ew_out -RedirectStandardError $__ew_err; "
                + "if (Test-Path $__ew_out) { Get-Content -Raw -ErrorAction SilentlyContinue $__ew_out }; "
                + "if (Test-Path $__ew_err) { Get-Content -Raw -ErrorAction SilentlyContinue $__ew_err | Write-Error -ErrorAction Continue }; "
                + "Write-Output ('" + marker + ":' + [int]$__ew_p.ExitCode)");

            var result = new TerminalCommandResult { Cwd = Cwd };
            var sw = Stopwatch.StartNew();
            int timeoutMs = Math.Max(1, timeoutSec) * 1000;
            string markerLine = null;

            try
            {
                while (sw.ElapsedMilliseconds < timeoutMs)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    markerLine = FindMarker(marker);
                    if (markerLine != null)
                    {
                        break;
                    }

                    await Task.Delay(80, cancellationToken).ConfigureAwait(false);
                }
            }
            catch (OperationCanceledException)
            {
                ProcessTreeKiller.KillDescendants(_process.Id);
                result.Cancelled = true;
                result.ExitCode = -1;
            }

            if (result.Cancelled)
            {
                FillOutput(result, marker);
                return result;
            }

            if (markerLine == null)
            {
                ProcessTreeKiller.KillDescendants(_process.Id);
                result.TimedOut = true;
                result.ExitCode = -1;
                FillOutput(result, marker);
                return result;
            }

            result.ExitCode = ParseExit(markerLine);
            FillOutput(result, marker);
            TryDelete(outFile);
            TryDelete(errFile);
            return result;
        }

        public void KillCurrentCommand()
        {
            if (_process != null && !_process.HasExited)
            {
                ProcessTreeKiller.KillDescendants(_process.Id);
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            try
            {
                if (_process != null && !_process.HasExited)
                {
                    ProcessTreeKiller.KillTree(_process.Id);
                }
            }
            catch
            {
            }

            try { _stdin?.Dispose(); } catch { }
            try { _process?.Dispose(); } catch { }
        }

        private string FindMarker(string marker)
        {
            lock (_gate)
            {
                string text = _stdoutBuf.ToString();
                var match = Regex.Match(text, marker + @":(-?\d+)");
                return match.Success ? match.Value : null;
            }
        }

        private static int ParseExit(string markerLine)
        {
            int idx = markerLine.LastIndexOf(':');
            if (idx < 0)
            {
                return 0;
            }

            int.TryParse(markerLine.Substring(idx + 1), out int code);
            return code;
        }

        private void FillOutput(TerminalCommandResult result, string marker)
        {
            string rawOut;
            string rawErr;
            lock (_gate)
            {
                rawOut = _stdoutBuf.ToString();
                rawErr = _stderrBuf.ToString();
            }

            rawOut = Regex.Replace(rawOut, marker + @":-?\d+\r?\n?", "");
            result.StdoutChars = rawOut.Length;
            result.StderrChars = rawErr.Length;
            result.Stdout = Truncate(rawOut, out bool t1);
            result.Stderr = Truncate(rawErr, out bool t2);
            result.Truncated = t1 || t2;
        }

        private static string Truncate(string text, out bool truncated)
        {
            if (text == null)
            {
                truncated = false;
                return "";
            }

            if (text.Length <= MaxOutputChars)
            {
                truncated = false;
                return text;
            }

            truncated = true;
            return text.Substring(0, MaxOutputChars);
        }

        private static string EscapePs(string value)
        {
            return (value ?? "").Replace("'", "''");
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }
    }
}
