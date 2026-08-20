using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// Desktop 侧 Named Pipe 服务：接收 YiWriteBrowserBridge（由 Chrome/Edge 拉起）转发的 NM 消息。
    /// </summary>
    public static class BrowserBridgePipeServer
    {
        public const string PipeName = "YiWrite.BrowserBridge";

        private static readonly object Gate = new object();
        private static CancellationTokenSource _cts;
        private static Task _acceptLoop;
        private static readonly List<ClientSession> Clients = new List<ClientSession>();

        public static event Action<JObject, ClientSession> MessageReceived;

        public static void Start()
        {
            lock (Gate)
            {
                if (_cts != null)
                {
                    return;
                }

                _cts = new CancellationTokenSource();
                CancellationToken ct = _cts.Token;
                _acceptLoop = Task.Run(() => AcceptLoopAsync(ct));
            }
        }

        public static void Stop()
        {
            CancellationTokenSource cts;
            lock (Gate)
            {
                cts = _cts;
                _cts = null;
            }

            try
            {
                cts?.Cancel();
            }
            catch
            {
                /* ignore */
            }

            lock (Gate)
            {
                foreach (ClientSession c in Clients.ToArray())
                {
                    c.Dispose();
                }

                Clients.Clear();
            }
        }

        public static IReadOnlyList<ClientSession> SnapshotClients()
        {
            lock (Gate)
            {
                return Clients.ToArray();
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
                    var session = new ClientSession(server);
                    lock (Gate)
                    {
                        Clients.Add(session);
                    }

                    _ = Task.Run(() => session.ReadLoopAsync(ct));
                    server = null;
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[BrowserBridgePipe] accept: " + ex.Message);
                    await Task.Delay(500, CancellationToken.None).ConfigureAwait(false);
                }
                finally
                {
                    try
                    {
                        server?.Dispose();
                    }
                    catch
                    {
                        /* ignore */
                    }
                }
            }
        }

        internal static void RemoveClient(ClientSession session)
        {
            lock (Gate)
            {
                Clients.Remove(session);
            }

            try
            {
                ExtensionHost.NotifyBridgeDisconnected(session);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[BrowserBridgePipe] disconnect notify: " + ex.Message);
            }
        }

        public sealed class ClientSession : IDisposable
        {
            private readonly NamedPipeServerStream _pipe;
            private readonly object _writeGate = new object();
            private bool _disposed;

            public ClientSession(NamedPipeServerStream pipe)
            {
                _pipe = pipe ?? throw new ArgumentNullException(nameof(pipe));
                SessionId = Guid.NewGuid().ToString("N");
            }

            public string SessionId { get; }

            public string BrowserKind { get; set; }

            public string ExtensionId { get; set; }

            public async Task ReadLoopAsync(CancellationToken ct)
            {
                try
                {
                    while (!ct.IsCancellationRequested && _pipe.IsConnected)
                    {
                        byte[] body = await ReadFrameAsync(_pipe, ct).ConfigureAwait(false);
                        if (body == null)
                        {
                            break;
                        }

                        string json = Encoding.UTF8.GetString(body);
                        JObject msg;
                        try
                        {
                            msg = JObject.Parse(json);
                        }
                        catch
                        {
                            continue;
                        }

                        MessageReceived?.Invoke(msg, this);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[BrowserBridgePipe] read: " + ex.Message);
                }
                finally
                {
                    RemoveClient(this);
                    Dispose();
                }
            }

            public Task SendAsync(JObject msg, CancellationToken ct = default(CancellationToken))
            {
                if (msg == null)
                {
                    throw new ArgumentNullException(nameof(msg));
                }

                byte[] body = Encoding.UTF8.GetBytes(msg.ToString(Formatting.None));
                return Task.Run(() =>
                {
                    lock (_writeGate)
                    {
                        if (_disposed || !_pipe.IsConnected)
                        {
                            throw new InvalidOperationException("bridge session closed");
                        }

                        WriteFrame(_pipe, body);
                    }
                }, ct);
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
                    _pipe.Dispose();
                }
                catch
                {
                    /* ignore */
                }
            }
        }

        private static async Task<byte[]> ReadFrameAsync(Stream stream, CancellationToken ct)
        {
            byte[] lenBuf = new byte[4];
            if (!await ReadExactAsync(stream, lenBuf, 4, ct).ConfigureAwait(false))
            {
                return null;
            }

            int len = BitConverter.ToInt32(lenBuf, 0);
            if (len <= 0 || len > 8 * 1024 * 1024)
            {
                throw new InvalidOperationException("bad frame len " + len);
            }

            byte[] body = new byte[len];
            if (!await ReadExactAsync(stream, body, len, ct).ConfigureAwait(false))
            {
                return null;
            }

            return body;
        }

        private static void WriteFrame(Stream stream, byte[] body)
        {
            byte[] lenBuf = BitConverter.GetBytes(body.Length);
            stream.Write(lenBuf, 0, 4);
            stream.Write(body, 0, body.Length);
            stream.Flush();
        }

        private static async Task<bool> ReadExactAsync(Stream stream, byte[] buffer, int count, CancellationToken ct)
        {
            int off = 0;
            while (off < count)
            {
                int n = await stream.ReadAsync(buffer, off, count - off, ct).ConfigureAwait(false);
                if (n <= 0)
                {
                    return false;
                }

                off += n;
            }

            return true;
        }
    }
}
