using System;

namespace WordAddIn1.PresentationHost
{
    /// <summary>约定 HTML 几何：每一层相对上一层（%），COM 用幻灯片点。</summary>
    internal static class PptHtmlGeom
    {
        public static void ChildPctToParentPct(
            double childLeft,
            double childTop,
            double childWidth,
            double childHeight,
            double parentLeft,
            double parentTop,
            double parentWidth,
            double parentHeight,
            out double outLeft,
            out double outTop,
            out double outWidth,
            out double outHeight)
        {
            outLeft = parentLeft + childLeft / 100.0 * parentWidth;
            outTop = parentTop + childTop / 100.0 * parentHeight;
            outWidth = childWidth / 100.0 * parentWidth;
            outHeight = childHeight / 100.0 * parentHeight;
        }

        public static void SlideBoxToParentPct(
            double slideLeft,
            double slideTop,
            double slideWidth,
            double slideHeight,
            double parentLeft,
            double parentTop,
            double parentWidth,
            double parentHeight,
            out double pctLeft,
            out double pctTop,
            out double pctWidth,
            out double pctHeight)
        {
            if (parentWidth <= 0.0001 || parentHeight <= 0.0001)
            {
                pctLeft = pctTop = pctWidth = pctHeight = 0;
                return;
            }

            pctLeft = (slideLeft - parentLeft) / parentWidth * 100.0;
            pctTop = (slideTop - parentTop) / parentHeight * 100.0;
            pctWidth = slideWidth / parentWidth * 100.0;
            pctHeight = slideHeight / parentHeight * 100.0;
        }

        public static string StyleFromSlideBox(
            double slideLeft,
            double slideTop,
            double slideWidth,
            double slideHeight,
            double parentLeft,
            double parentTop,
            double parentWidth,
            double parentHeight)
        {
            SlideBoxToParentPct(
                slideLeft,
                slideTop,
                slideWidth,
                slideHeight,
                parentLeft,
                parentTop,
                parentWidth,
                parentHeight,
                out double l,
                out double t,
                out double w,
                out double h);
            return PptConventionHtml.BuildStyle(l, t, w, h);
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
