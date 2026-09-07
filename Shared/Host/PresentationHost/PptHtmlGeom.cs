using System;

namespace WordAddIn1.PresentationHost
{
    /// <summary>约定 HTML 几何：% 一律相对幻灯片；COM 用幻灯片点。</summary>
    internal static class PptHtmlGeom
    {
        public static string StyleFromSlidePoints(
            double left,
            double top,
            double width,
            double height,
            double slideWidth,
            double slideHeight)
        {
            if (slideWidth <= 0.0001 || slideHeight <= 0.0001)
            {
                return PptConventionHtml.BuildStyle(0, 0, 0, 0);
            }

            return PptConventionHtml.BuildStyle(
                left / slideWidth * 100.0,
                top / slideHeight * 100.0,
                width / slideWidth * 100.0,
                height / slideHeight * 100.0);
        }

        /// <summary>
        /// 把已是页 % 的框，从源页框仿射到目标页框（extends：组件根原页框 → 实例页框）。
        /// </summary>
        public static void MapSlidePctBox(
            double srcLeft,
            double srcTop,
            double srcWidth,
            double srcHeight,
            double fromLeft,
            double fromTop,
            double fromWidth,
            double fromHeight,
            double toLeft,
            double toTop,
            double toWidth,
            double toHeight,
            out double outLeft,
            out double outTop,
            out double outWidth,
            out double outHeight)
        {
            if (fromWidth <= 0.0001 || fromHeight <= 0.0001)
            {
                outLeft = toLeft;
                outTop = toTop;
                outWidth = toWidth;
                outHeight = toHeight;
                return;
            }

            outLeft = toLeft + (srcLeft - fromLeft) / fromWidth * toWidth;
            outTop = toTop + (srcTop - fromTop) / fromHeight * toHeight;
            outWidth = srcWidth / fromWidth * toWidth;
            outHeight = srcHeight / fromHeight * toHeight;
        }

        public static int CountNodes(PptHtmlShapeNode node)
        {
            if (node == null)
            {
                return 0;
            }

            int n = 1;
            if (node.Children != null)
            {
                foreach (PptHtmlShapeNode child in node.Children)
                {
                    n += CountNodes(child);
                }
            }

            return n;
        }

        public static int CountNodes(System.Collections.Generic.IList<PptHtmlShapeNode> nodes)
        {
            if (nodes == null)
            {
                return 0;
            }

            int n = 0;
            foreach (PptHtmlShapeNode node in nodes)
            {
                n += CountNodes(node);
            }

            return n;
        }

        public static int CountApplyNodes(PptHtmlApplyNode node)
        {
            if (node == null)
            {
                return 0;
            }

            int n = 1;
            if (node.Children != null)
            {
                foreach (PptHtmlApplyNode child in node.Children)
                {
                    n += CountApplyNodes(child);
                }
            }

            return n;
        }

        public static int CountApplyNodes(System.Collections.Generic.IList<PptHtmlApplyNode> nodes)
        {
            if (nodes == null)
            {
                return 0;
            }

            int n = 0;
            foreach (PptHtmlApplyNode node in nodes)
            {
                n += CountApplyNodes(node);
            }

            return n;
        }
    }
}
