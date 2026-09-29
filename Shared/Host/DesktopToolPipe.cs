using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// Desktop 侧工具调用管道：测试短套把 F_* 打进正在跑的易写进程（与扩展同一根桥）。
    /// </summary>
    public static class DesktopToolPipe
    {
        public const string PipeName = "YiWrite.DesktopTools";

        private static readonly object Gate = new object();
        private static CancellationTokenSource _cts;
        private static Func<string, Dictionary<string, object>, Task<ToolResult>> _invoke;

        public static void Start(Func<string, Dictionary<string, object>, Task<ToolResult>> invoke)
        {
            if (invoke == null)
            {
                throw new ArgumentNullException(nameof(invoke));
            }

            lock (Gate)
            {
                _invoke = invoke;
                if (_cts != null)
                {
                    return;
                }

                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                Task.Run(() => AcceptLoopAsync(ct));
            }
        }

        public static void Stop()
        {
            CancellationTokenSource cts;
            lock (Gate)
            {
                cts = _cts;
                _cts = null;
                _invoke = null;
            }

            try
            {
                cts?.Cancel();
            }
            catch
            {
            }
        }

        public static bool TryWaitUntilAvailable(int timeoutMs)
        {
            DateTime deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs < 0 ? 0 : timeoutMs);
            while (DateTime.UtcNow < deadline)
            {
                if (TryPing())
                {
                    return true;
                }

                Thread.Sleep(200);
            }

            return TryPing();
        }

        public static ToolResult Call(string tool, Dictionary<string, object> args, int timeoutMs = 120000)
        {
            var req = new JObject
            {
                ["tool"] = tool ?? "",
                ["args"] = args == null ? new JObject() : JObject.FromObject(args)
            };
            try
            {
                string raw = Transact(req.ToString(Newtonsoft.Json.Formatting.None), timeoutMs);
                JObject jo = JObject.Parse(raw);
                var dataTok = jo["data"];
                Dictionary<string, object> data = null;
                if (dataTok != null && dataTok.Type == JTokenType.Object)
                {
                    data = dataTok.ToObject<Dictionary<string, object>>();
                }

                return new ToolResult
                {
                    Success = jo.Value<bool?>("success") ?? false,
                    Error = jo.Value<string>("error"),
                    Data = data
                };
            }
            catch (Exception ex)
            {
                return new ToolResult
                {
                    Success = false,
                    Error = "连不上易写 Desktop（" + PipeName + "）: " + ex.Message
                };
            }
        }

        private static bool TryPing()
        {
            try
            {
                using (var pipe = new NamedPipeClientStream(
                    ".",
                    PipeName,
                    PipeDirection.InOut,
                    PipeOptions.None))
                {
                    pipe.Connect(250);
                    return pipe.IsConnected;
                }
            }
            catch
            {
                return false;
            }
        }

        private static string Transact(string json, int timeoutMs)
        {
            using (var cts = new CancellationTokenSource(timeoutMs))
            using (var pipe = new NamedPipeClientStream(
                ".",
                PipeName,
                PipeDirection.InOut,
                PipeOptions.Asynchronous))
            {
                pipe.Connect(Math.Min(4000, Math.Max(500, timeoutMs)));
                WriteFrame(pipe, json, cts.Token);
                return ReadFrame(pipe, cts.Token);
            }
        }

        private static async Task AcceptLoopAsync(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                NamedPipeServerStream server = null;
                try
                {
                    server = new NamedPipeServerStream(
                        PipeName,
                        PipeDirection.InOut,
                        NamedPipeServerStream.MaxAllowedServerInstances,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous);
                    await server.WaitForConnectionAsync(ct).ConfigureAwait(false);
                    NamedPipeServerStream session = server;
                    server = null;
                    _ = Task.Run(() => ServeOneAsync(session, ct));
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception)
                {
                    await Task.Delay(300, CancellationToken.None).ConfigureAwait(false);
                }
                finally
                {
                    try
                    {
                        server?.Dispose();
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static async Task ServeOneAsync(NamedPipeServerStream pipe, CancellationToken ct)
        {
            try
            {
                using (pipe)
                {
                    string raw = await ReadFrameAsync(pipe, ct).ConfigureAwait(false);
                    JObject req = JObject.Parse(raw);
                    string tool = (req.Value<string>("tool") ?? "").Trim();
                    Dictionary<string, object> args = null;
                    JToken argsTok = req["args"];
                    if (argsTok != null && argsTok.Type == JTokenType.Object)
                    {
                        args = argsTok.ToObject<Dictionary<string, object>>();
                    }

                    Func<string, Dictionary<string, object>, Task<ToolResult>> invoke;
                    lock (Gate)
                    {
                        invoke = _invoke;
                    }

                    ToolResult result;
                    if (invoke == null)
                    {
                        result = new ToolResult { Success = false, Error = "Desktop 工具未就绪" };
                    }
                    else
                    {
                        result = await invoke(tool, args ?? new Dictionary<string, object>())
                            .ConfigureAwait(false);
                    }

                    var data = result != null && result.Data != null
                        ? JToken.FromObject(result.Data)
                        : JValue.CreateNull();
                    var resp = new JObject
                    {
                        ["success"] = result != null && result.Success,
                        ["error"] = result != null ? result.Error : "空回包",
                        ["data"] = data
                    };
                    await WriteFrameAsync(pipe, resp.ToString(Newtonsoft.Json.Formatting.None), ct)
                        .ConfigureAwait(false);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void WriteFrame(Stream s, string json, CancellationToken ct)
        {
            WriteFrameAsync(s, json, ct).GetAwaiter().GetResult();
        }

        private static string ReadFrame(Stream s, CancellationToken ct)
        {
            return ReadFrameAsync(s, ct).GetAwaiter().GetResult();
        }

        private static async Task WriteFrameAsync(Stream s, string json, CancellationToken ct)
        {
            byte[] payload = Encoding.UTF8.GetBytes(json ?? "");
            byte[] len = BitConverter.GetBytes(payload.Length);
            await s.WriteAsync(len, 0, 4, ct).ConfigureAwait(false);
            await s.WriteAsync(payload, 0, payload.Length, ct).ConfigureAwait(false);
            await s.FlushAsync(ct).ConfigureAwait(false);
        }

        private static async Task<string> ReadFrameAsync(Stream s, CancellationToken ct)
        {
            byte[] lenBuf = new byte[4];
            await ReadExactAsync(s, lenBuf, ct).ConfigureAwait(false);
            int n = BitConverter.ToInt32(lenBuf, 0);
            if (n < 0 || n > 8 * 1024 * 1024)
            {
                throw new InvalidOperationException("Desktop 工具管道帧过大");
            }

            byte[] payload = new byte[n];
            if (n > 0)
            {
                await ReadExactAsync(s, payload, ct).ConfigureAwait(false);
            }

            return Encoding.UTF8.GetString(payload);
        }

        private static async Task ReadExactAsync(Stream s, byte[] buf, CancellationToken ct)
        {
            int off = 0;
            while (off < buf.Length)
            {
                int n = await s.ReadAsync(buf, off, buf.Length - off, ct).ConfigureAwait(false);
                if (n <= 0)
                {
                    throw new EndOfStreamException("Desktop 工具管道断开");
                }

                off += n;
            }
        }
    }
}
