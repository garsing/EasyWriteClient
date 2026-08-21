using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WordAddIn1.Terminal
{
    internal static class ManagedPythonRuntime
    {
        public const string Version = "3.12.10";
        public const string DefaultZipUrl =
            "https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip";
        public const string DefaultSha256 =
            "4acbed6dd1c744b0376e3b1cf57ce906f9dc9e95e68824584c8099a63025a3c3";
        private const string GetPipUrl = "https://bootstrap.pypa.io/get-pip.py";

        private static readonly SemaphoreSlim InstallLock = new SemaphoreSlim(1, 1);

        public static string RootDirectory
        {
            get
            {
                string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                return Path.Combine(local, "EasyWriteDesktop", "binaries", "python");
            }
        }

        public static string VersionDirectory => Path.Combine(RootDirectory, "versions", Version);

        public static string PythonExe => Path.Combine(VersionDirectory, "python.exe");

        public static string SiteDirectory => Path.Combine(RootDirectory, "envs", "default");

        public static bool IsInstalled()
        {
            return File.Exists(PythonExe)
                && File.Exists(Path.Combine(VersionDirectory, "python312.zip"))
                && Directory.Exists(Path.Combine(SiteDirectory, "docx"));
        }

        public static bool IsHealthy()
        {
            if (!IsInstalled())
            {
                return false;
            }

            RepairPth(VersionDirectory);
            return Probe("import encodings, docx");
        }

        public static async Task EnsureInstalledAsync(CancellationToken cancellationToken)
        {
            if (IsHealthy())
            {
                return;
            }

            await InstallLock.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                if (IsHealthy())
                {
                    return;
                }

                Directory.CreateDirectory(RootDirectory);
                string lockPath = Path.Combine(RootDirectory, ".install.lock");
                using (var lockStream = new FileStream(
                    lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    await InstallCoreAsync(cancellationToken).ConfigureAwait(false);
                    if (!IsHealthy())
                    {
                        throw new InvalidOperationException(
                            "托管 Python 安装失败：安装完成后无法 import encodings / docx");
                    }
                }
            }
            finally
            {
                InstallLock.Release();
            }
        }

        private static async Task InstallCoreAsync(CancellationToken cancellationToken)
        {
            if (Directory.Exists(VersionDirectory))
            {
                Directory.Delete(VersionDirectory, true);
            }

            Directory.CreateDirectory(VersionDirectory);
            Directory.CreateDirectory(SiteDirectory);

            string zipPath = Path.Combine(RootDirectory, "python-" + Version + "-embed-amd64.zip");
            await DownloadAsync(DefaultZipUrl, zipPath, cancellationToken).ConfigureAwait(false);
            VerifySha256(zipPath, DefaultSha256);
            ZipFile.ExtractToDirectory(zipPath, VersionDirectory);
            if (!File.Exists(PythonExe) || !File.Exists(Path.Combine(VersionDirectory, "python312.zip")))
            {
                throw new InvalidOperationException("托管 Python 安装失败：解压后缺少 python.exe 或 python312.zip");
            }

            RepairPth(VersionDirectory);
            string getPip = Path.Combine(RootDirectory, "get-pip.py");
            await DownloadAsync(GetPipUrl, getPip, cancellationToken).ConfigureAwait(false);

            RunPython("\"" + getPip + "\"", VersionDirectory);
            RunPython("-m pip install --target \"" + SiteDirectory + "\" python-docx", VersionDirectory);

            File.WriteAllText(
                Path.Combine(RootDirectory, "state.json"),
                "{\"ver\":\"" + Version + "\",\"installed_at\":\"" + DateTime.UtcNow.ToString("o") + "\"}",
                Encoding.UTF8);
        }

        /// <summary>
        /// embed 的 ._pth 必须无 BOM。Encoding.UTF8 默认带 BOM，会导致找不到 python312.zip，
        /// 进而报 No module named encodings。
        /// </summary>
        internal static void RepairPth(string versionDir)
        {
            if (!Directory.Exists(versionDir))
            {
                return;
            }

            foreach (string pth in Directory.GetFiles(versionDir, "python*._pth"))
            {
                string text = File.ReadAllText(pth, Encoding.UTF8);
                if (text.Length > 0 && text[0] == '\uFEFF')
                {
                    text = text.Substring(1);
                }

                text = text.Replace("#import site", "import site");
                if (text.IndexOf("import site", StringComparison.Ordinal) < 0)
                {
                    text = text.TrimEnd() + Environment.NewLine + "import site" + Environment.NewLine;
                }

                string siteLine = SiteDirectory.Replace('\\', '/');
                if (text.IndexOf(siteLine, StringComparison.OrdinalIgnoreCase) < 0
                    && text.IndexOf(SiteDirectory, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    int siteIdx = text.IndexOf("import site", StringComparison.Ordinal);
                    if (siteIdx >= 0)
                    {
                        text = text.Insert(siteIdx, siteLine + Environment.NewLine);
                    }
                    else
                    {
                        text = text.TrimEnd() + Environment.NewLine + siteLine + Environment.NewLine;
                    }
                }

                File.WriteAllText(pth, text, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            }
        }

        private static bool Probe(string code)
        {
            try
            {
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = PythonExe,
                    Arguments = "-c \"" + code + "\"",
                    WorkingDirectory = VersionDirectory,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };
                psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
                psi.EnvironmentVariables["PYTHONUTF8"] = "1";
                psi.EnvironmentVariables.Remove("PYTHONPATH");
                psi.EnvironmentVariables.Remove("PYTHONHOME");
                using (var p = System.Diagnostics.Process.Start(psi))
                {
                    if (p == null)
                    {
                        return false;
                    }

                    p.WaitForExit(15000);
                    return p.HasExited && p.ExitCode == 0;
                }
            }
            catch
            {
                return false;
            }
        }

        private static async Task DownloadAsync(string url, string dest, CancellationToken cancellationToken)
        {
            using (var client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMinutes(5);
                using (var response = await client.GetAsync(url, cancellationToken).ConfigureAwait(false))
                {
                    response.EnsureSuccessStatusCode();
                    using (Stream src = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                    using (var dst = File.Create(dest))
                    {
                        await src.CopyToAsync(dst).ConfigureAwait(false);
                    }
                }
            }
        }

        private static void VerifySha256(string path, string expected)
        {
            using (var sha = SHA256.Create())
            using (var stream = File.OpenRead(path))
            {
                byte[] hash = sha.ComputeHash(stream);
                var sb = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash)
                {
                    sb.Append(b.ToString("x2"));
                }

                string actual = sb.ToString();
                if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "托管 Python 安装失败：zip SHA256 不匹配（expected=" + expected + " actual=" + actual + "）");
                }
            }
        }

        private static void RunPython(string args, string workingDir)
        {
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = PythonExe,
                Arguments = args,
                WorkingDirectory = workingDir,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
            psi.EnvironmentVariables["PYTHONUTF8"] = "1";
            psi.EnvironmentVariables.Remove("PYTHONPATH");
            psi.EnvironmentVariables.Remove("PYTHONHOME");
            using (var p = System.Diagnostics.Process.Start(psi))
            {
                if (p == null)
                {
                    throw new InvalidOperationException("托管 Python 安装失败：无法启动 python.exe");
                }

                p.WaitForExit();
                if (p.ExitCode != 0)
                {
                    string err = p.StandardError.ReadToEnd();
                    throw new InvalidOperationException("托管 Python 安装失败: " + err);
                }
            }
        }
    }
}
