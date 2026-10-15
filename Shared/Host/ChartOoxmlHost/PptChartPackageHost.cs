using System;
using System.Collections.Generic;
using System.IO.Compression;
using WordAddIn1.PresentationHost;

namespace WordAddIn1.ChartOoxmlHost
{
    /// <summary>PPT/WPP 图表改包入口。内部委托 PptHtmlOoxmlIo，不抄瘦包。</summary>
    internal static class PptChartPackageHost
    {
        public static bool TrySaveCopy(object presentation, string path, out string error)
        {
            return PptHtmlOoxmlIo.TrySaveCopy(presentation, path, out error);
        }

        public static string TrySaveCopyTemp(object presentation, string tag, out string error)
        {
            return PptHtmlOoxmlIo.TrySaveCopyTemp(presentation, tag, out error);
        }

        public static bool TryReadSavedPackage(
            object presentation,
            Func<string, bool> read,
            out string error,
            string tag = "读包")
        {
            return PptHtmlOoxmlIo.TryReadSavedPackage(presentation, read, out error, tag);
        }

        public static bool TryRewriteAndCopyBack(
            object destPres,
            int destSlideIndex,
            Func<ZipArchive, IDictionary<string, byte[]>, bool> rewrite,
            Func<object, int, object> findOnCopy,
            Func<bool> pasteToDest,
            out string error,
            string alreadySavedPath = null,
            string tag = "改包",
            List<string> warnings = null)
        {
            return PptHtmlOoxmlIo.TryRewriteAndCopyBack(
                destPres,
                destSlideIndex,
                rewrite,
                findOnCopy,
                pasteToDest,
                out error,
                alreadySavedPath,
                tag,
                warnings);
        }

        public static object TryOpenCopy(object app, string path, out string error, bool readOnly = false)
        {
            return PptHtmlOoxmlIo.TryOpenCopy(app, path, out error, readOnly);
        }

        public static object TryFindShapeById(object presentation, int slideIndex, int shapeId)
        {
            return PptHtmlOoxmlIo.TryFindShapeById(presentation, slideIndex, shapeId);
        }

        public static void TryCloseCopy(object presentation)
        {
            PptHtmlOoxmlIo.TryCloseCopy(presentation);
        }

        public static void TryDeleteFile(string path)
        {
            PptHtmlOoxmlIo.TryDeleteFile(path);
        }

        public static string NewTempPath(string tag)
        {
            return PptHtmlOoxmlIo.NewTempPath(tag);
        }
    }
}
