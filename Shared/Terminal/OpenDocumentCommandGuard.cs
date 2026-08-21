using System;
using System.IO;

namespace WordAddIn1.Terminal
{
    internal static class OpenDocumentCommandGuard
    {
        public static bool TryFindBlockedPath(string command, string cwd, out string blockedPath)
        {
            blockedPath = null;
            string haystack = (command ?? "") + " " + (cwd ?? "");
            if (string.IsNullOrWhiteSpace(haystack))
            {
                return false;
            }

            foreach (IOperationChannel channel in ChannelRegistry.Snapshot())
            {
                if (!IsOffice(channel.Kind))
                {
                    continue;
                }

                string path = TryGetFilePath(channel);
                if (string.IsNullOrWhiteSpace(path))
                {
                    continue;
                }

                string full;
                try
                {
                    full = Path.GetFullPath(path);
                }
                catch
                {
                    full = path;
                }

                if (haystack.IndexOf(full, StringComparison.OrdinalIgnoreCase) >= 0
                    || haystack.IndexOf(path, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    blockedPath = full;
                    return true;
                }
            }

            return false;
        }

        private static bool IsOffice(ChannelKind kind)
        {
            return kind == ChannelKind.Word
                || kind == ChannelKind.Wps
                || kind == ChannelKind.Excel
                || kind == ChannelKind.Et
                || kind == ChannelKind.Ppt
                || kind == ChannelKind.Wpp;
        }

        private static string TryGetFilePath(IOperationChannel channel)
        {
            if (channel is WordChannel word)
            {
                return word.FilePath;
            }

            if (channel is WpsChannel wps)
            {
                return wps.FilePath;
            }

            if (channel is ExcelChannel excel)
            {
                return excel.FilePath;
            }

            if (channel is EtChannel et)
            {
                return et.FilePath;
            }

            if (channel is PptChannel ppt)
            {
                return ppt.FilePath;
            }

            if (channel is WppChannel wpp)
            {
                return wpp.FilePath;
            }

            return null;
        }
    }
}
