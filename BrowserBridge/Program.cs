using System;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace YiWriteBrowserBridge
{
    /// <summary>
    /// Chrome/Edge Native Messaging 宿主：stdin/stdout ↔ Desktop Named Pipe。
    /// 由浏览器启动；易写 Desktop 须先监听管道。
    /// </summary>
    internal static class Program
    {
        public const string PipeName = "YiWrite.BrowserBridge";

        [STAThread]
        private static void Main()
        {
            try
            {
                RunAsync().GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                TryLog("fatal: " + ex);
            }
        }

        private static async Task RunAsync()
        {
            using (var cts = new CancellationTokenSource())
            {
                NamedPipeClientStream pipe = null;
                for (int i = 0; i < 40; i++)
                {
                    try
                    {
                        pipe = new NamedPipeClientStream(
                            ".",
                            PipeName,
                            PipeDirection.InOut,
                            PipeOptions.Asynchronous);
                        await pipe.ConnectAsync(500).ConfigureAwait(false);
                        break;
                    }
                    catch
                    {
                        pipe?.Dispose();
                        pipe = null;
                        await Task.Delay(250).ConfigureAwait(false);
                    }
                }

                if (pipe == null || !pipe.IsConnected)
                {
                    TryLog("Desktop pipe not available: " + PipeName);
                    // 仍消费 stdin，避免 Chrome 报错刷屏；回复错误 hello 无法送达扩展侧以外
                    return;
                }

                var stdin = Console.OpenStandardInput();
                var stdout = Console.OpenStandardOutput();

                var nmToPipe = Task.Run(() => ForwardNmToPipe(stdin, pipe, cts.Token));
                var pipeToNm = Task.Run(() => ForwardPipeToNm(pipe, stdout, cts.Token));

                await Task.WhenAny(nmToPipe, pipeToNm).ConfigureAwait(false);
                cts.Cancel();
                try { pipe.Dispose(); } catch { /* ignore */ }
            }
        }

        private static async Task ForwardNmToPipe(
            Stream stdin,
            NamedPipeClientStream pipe,
            CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                byte[] frame = await ReadLengthPrefixedAsync(stdin, ct).ConfigureAwait(false);
                if (frame == null)
                {
                    break;
                }

                await WriteLengthPrefixedAsync(pipe, frame, ct).ConfigureAwait(false);
            }
        }

        private static async Task ForwardPipeToNm(
            NamedPipeClientStream pipe,
            Stream stdout,
            CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                byte[] frame = await ReadLengthPrefixedAsync(pipe, ct).ConfigureAwait(false);
                if (frame == null)
                {
                    break;
                }

                await WriteLengthPrefixedAsync(stdout, frame, ct).ConfigureAwait(false);
                await stdout.FlushAsync(ct).ConfigureAwait(false);
            }
        }

        private static async Task<byte[]> ReadLengthPrefixedAsync(Stream stream, CancellationToken ct)
        {
            byte[] lenBuf = new byte[4];
            if (!await ReadExactAsync(stream, lenBuf, 4, ct).ConfigureAwait(false))
            {
                return null;
            }

            int len = BitConverter.ToInt32(lenBuf, 0);
            if (len <= 0 || len > 8 * 1024 * 1024)
            {
                throw new InvalidOperationException("invalid frame length: " + len);
            }

            byte[] body = new byte[len];
            if (!await ReadExactAsync(stream, body, len, ct).ConfigureAwait(false))
            {
                return null;
            }

            return body;
        }

        private static async Task WriteLengthPrefixedAsync(Stream stream, byte[] body, CancellationToken ct)
        {
            byte[] lenBuf = BitConverter.GetBytes(body.Length);
            await stream.WriteAsync(lenBuf, 0, 4, ct).ConfigureAwait(false);
            await stream.WriteAsync(body, 0, body.Length, ct).ConfigureAwait(false);
        }

        private static async Task<bool> ReadExactAsync(
            Stream stream,
            byte[] buffer,
            int count,
            CancellationToken ct)
        {
            int offset = 0;
            while (offset < count)
            {
                int n = await stream.ReadAsync(buffer, offset, count - offset, ct).ConfigureAwait(false);
                if (n <= 0)
                {
                    return false;
                }

                offset += n;
            }

            return true;
        }

        private static void TryLog(string line)
        {
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "YiWrite",
                    "logs");
                Directory.CreateDirectory(dir);
                File.AppendAllText(
                    Path.Combine(dir, "browser-bridge.log"),
                    DateTime.Now.ToString("s") + " " + line + Environment.NewLine,
                    Encoding.UTF8);
            }
            catch
            {
                /* ignore */
            }
        }
    }
}
