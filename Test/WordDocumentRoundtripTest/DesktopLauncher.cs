using System;
using System.Diagnostics;
using System.IO;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 测试程序自己拉起易写 Desktop：管道不在就先构建（缺 exe 时）再启动，然后等管道。
    /// 不 Quit 已有 Desktop；已在跑就直接复用。
    /// </summary>
    internal static class DesktopLauncher
    {
        /// <summary>确保 Desktop 管道可用。返回 true=可用。会把过程打到控制台（测试进度可见）。</summary>
        public static bool EnsureAvailable()
        {
            if (DesktopToolPipe.TryWaitUntilAvailable(1500))
            {
                Console.WriteLine("[desktop] 管道已在，复用现有 Desktop");
                return true;
            }

            string exe = FindDesktopExe();
            if (exe == null)
            {
                Console.WriteLine("[desktop] 未找到 Desktop exe，先构建 EasyWriteClient.Desktop ...");
                if (!BuildDesktop())
                {
                    Console.WriteLine("[desktop] 构建失败");
                    return false;
                }

                exe = FindDesktopExe();
                if (exe == null)
                {
                    Console.WriteLine("[desktop] 构建后仍找不到 exe");
                    return false;
                }
            }

            Console.WriteLine("[desktop] 启动: " + exe);
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = exe,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine("[desktop] 启动失败: " + ex.Message);
                return false;
            }

            Console.WriteLine("[desktop] 等管道 " + DesktopToolPipe.PipeName + " ...");
            if (!DesktopToolPipe.TryWaitUntilAvailable(60000))
            {
                Console.WriteLine("[desktop] 等管道超时（60s）");
                return false;
            }

            Console.WriteLine("[desktop] 管道就绪");
            return true;
        }

        private static string FindDesktopExe()
        {
            string root = RepoRoot();
            if (root == null)
            {
                return null;
            }

            string[] candidates =
            {
                Path.Combine(root, @"EasyWriteClient\Desktop\bin\Debug\net472\EasyWriteClient.Desktop.exe"),
                Path.Combine(root, @"EasyWriteClient\Desktop\bin\Release\net472\EasyWriteClient.Desktop.exe")
            };
            foreach (string c in candidates)
            {
                if (File.Exists(c))
                {
                    return c;
                }
            }

            return null;
        }

        private static bool BuildDesktop()
        {
            string root = RepoRoot();
            if (root == null)
            {
                return false;
            }

            string proj = Path.Combine(root, @"EasyWriteClient\Desktop\EasyWriteClient.Desktop.csproj");
            if (!File.Exists(proj))
            {
                return false;
            }

            var psi = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = "build \"" + proj + "\" -c Debug --nologo",
                WorkingDirectory = root,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            try
            {
                using (Process p = Process.Start(psi))
                {
                    string stdout = p.StandardOutput.ReadToEnd();
                    string stderr = p.StandardError.ReadToEnd();
                    p.WaitForExit(300000);
                    string tail = Tail(stdout + stderr, 12);
                    Console.WriteLine(tail);
                    return p.ExitCode == 0;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("[desktop] 构建异常: " + ex.Message);
                return false;
            }
        }

        private static string Tail(string text, int lines)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string[] arr = text.Replace("\r\n", "\n").Split('\n');
            int start = Math.Max(0, arr.Length - lines);
            return string.Join(Environment.NewLine, arr, start, arr.Length - start);
        }

        /// <summary>从测试 exe 所在目录往上找含 EasyWriteClient 的仓根。</summary>
        private static string RepoRoot()
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            for (int i = 0; i < 8 && !string.IsNullOrEmpty(dir); i++)
            {
                if (Directory.Exists(Path.Combine(dir, "EasyWriteClient"))
                    && File.Exists(Path.Combine(dir, @"EasyWriteClient\Desktop\EasyWriteClient.Desktop.csproj")))
                {
                    return dir;
                }

                DirectoryInfo parent = Directory.GetParent(dir);
                dir = parent == null ? null : parent.FullName;
            }

            // 兜底：硬编码仓根（本机工作区）
            const string fallback = @"D:\project\EasyWrite";
            return File.Exists(Path.Combine(fallback, @"EasyWriteClient\Desktop\EasyWriteClient.Desktop.csproj"))
                ? fallback
                : null;
        }
    }
}
