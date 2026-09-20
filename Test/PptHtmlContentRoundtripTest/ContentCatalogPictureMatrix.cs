using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using WordAddIn1.PresentationHost;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>
    /// 图片属性交叉矩阵（pic-mx-*）。覆盖几何网格、源色、换图保框、换图改框、
    /// 旋转与负例；合计 >=100。
    /// </summary>
    internal static partial class ContentCatalog
    {
        private static void AddPictureMatrixCases(List<ContentCase> list, ref int page)
        {
            int before = list.Count;
            AddPicMxGeoCreates(list, ref page);
            AddPicMxSrcCreates(list, ref page);
            AddPicMxReplaceKeepGeo(list, ref page);
            AddPicMxReplaceChangeGeo(list, ref page);
            AddPicMxRotations(list, ref page);
            AddPicMxNegatives(list, ref page);
            int added = list.Count - before;
            if (added < 100)
            {
                throw new InvalidOperationException(
                    "图片矩阵用例不足 100（实际 " + added + "），请扩 ContentCatalogPictureMatrix。");
            }
        }

        private static void AddPicMxGeoCreates(List<ContentCase> list, ref int page)
        {
            // 40 组几何：覆盖角落、中带、宽扁、窄高等
            double[][] geos =
            {
                new[] { 5.0, 5.0, 20.0, 20.0 }, new[] { 10.0, 10.0, 30.0, 25.0 }, new[] { 15.0, 8.0, 40.0, 20.0 },
                new[] { 20.0, 15.0, 25.0, 35.0 }, new[] { 25.0, 20.0, 35.0, 30.0 }, new[] { 30.0, 10.0, 45.0, 25.0 },
                new[] { 35.0, 25.0, 20.0, 40.0 }, new[] { 40.0, 5.0, 30.0, 30.0 }, new[] { 45.0, 30.0, 25.0, 25.0 },
                new[] { 50.0, 15.0, 35.0, 35.0 }, new[] { 55.0, 40.0, 20.0, 30.0 }, new[] { 60.0, 10.0, 25.0, 40.0 },
                new[] { 8.0, 45.0, 40.0, 25.0 }, new[] { 12.0, 50.0, 30.0, 20.0 }, new[] { 18.0, 35.0, 50.0, 20.0 },
                new[] { 22.0, 55.0, 20.0, 25.0 }, new[] { 28.0, 42.0, 35.0, 28.0 }, new[] { 33.0, 12.0, 28.0, 45.0 },
                new[] { 38.0, 48.0, 22.0, 22.0 }, new[] { 42.0, 22.0, 40.0, 18.0 }, new[] { 48.0, 8.0, 18.0, 50.0 },
                new[] { 52.0, 28.0, 32.0, 32.0 }, new[] { 58.0, 18.0, 28.0, 38.0 }, new[] { 6.0, 60.0, 45.0, 18.0 },
                new[] { 14.0, 28.0, 22.0, 42.0 }, new[] { 24.0, 6.0, 48.0, 22.0 }, new[] { 34.0, 38.0, 26.0, 26.0 },
                new[] { 44.0, 52.0, 30.0, 20.0 }, new[] { 54.0, 6.0, 20.0, 45.0 }, new[] { 62.0, 35.0, 22.0, 30.0 },
                new[] { 7.0, 18.0, 55.0, 20.0 }, new[] { 16.0, 62.0, 35.0, 18.0 }, new[] { 26.0, 32.0, 18.0, 48.0 },
                new[] { 36.0, 58.0, 40.0, 18.0 }, new[] { 46.0, 14.0, 24.0, 36.0 }, new[] { 56.0, 46.0, 28.0, 24.0 },
                new[] { 9.0, 36.0, 32.0, 28.0 }, new[] { 19.0, 24.0, 38.0, 22.0 }, new[] { 29.0, 44.0, 42.0, 20.0 },
                new[] { 39.0, 16.0, 20.0, 50.0 }
            };

            for (int i = 0; i < geos.Length; i++)
            {
                double l = geos[i][0], t = geos[i][1], w = geos[i][2], h = geos[i][3];
                string name = "pic-mx-geo-" + i.ToString("00", CultureInfo.InvariantCulture)
                    + "-" + Tag(l) + "x" + Tag(t) + "-" + Tag(w) + "x" + Tag(h);
                double el = l, et = t, ew = w, eh = h;
                Add(list, ref page, name,
                    a => ContentHtml.PictureCreate(a.RedPng, el, et, ew, eh),
                    null,
                    ctx => GeoEq(PreferCreated(ctx, "picture"), el, et, ew, eh));
            }
        }

        private static void AddPicMxSrcCreates(List<ContentCase> list, ref int page)
        {
            var specs = new[]
            {
                new { Tag = "red", Pick = (Func<ContentAssets, string>)(a => a.RedPng), L = 10.0, T = 10.0, W = 30.0, H = 30.0 },
                new { Tag = "blue", Pick = (Func<ContentAssets, string>)(a => a.BluePng), L = 40.0, T = 10.0, W = 30.0, H = 30.0 },
                new { Tag = "green", Pick = (Func<ContentAssets, string>)(a => a.GreenPng), L = 20.0, T = 40.0, W = 35.0, H = 30.0 },
                new { Tag = "red-wide", Pick = (Func<ContentAssets, string>)(a => a.RedPng), L = 5.0, T = 55.0, W = 55.0, H = 20.0 },
                new { Tag = "blue-tall", Pick = (Func<ContentAssets, string>)(a => a.BluePng), L = 65.0, T = 5.0, W = 20.0, H = 55.0 },
                new { Tag = "green-sq", Pick = (Func<ContentAssets, string>)(a => a.GreenPng), L = 35.0, T = 25.0, W = 25.0, H = 25.0 },
                new { Tag = "red-sm", Pick = (Func<ContentAssets, string>)(a => a.RedPng), L = 70.0, T = 70.0, W = 15.0, H = 15.0 },
                new { Tag = "blue-mid", Pick = (Func<ContentAssets, string>)(a => a.BluePng), L = 15.0, T = 20.0, W = 40.0, H = 35.0 },
                new { Tag = "green-wide", Pick = (Func<ContentAssets, string>)(a => a.GreenPng), L = 8.0, T = 8.0, W = 60.0, H = 18.0 },
                new { Tag = "red-l", Pick = (Func<ContentAssets, string>)(a => a.RedPng), L = 50.0, T = 40.0, W = 35.0, H = 40.0 },
                new { Tag = "blue-l", Pick = (Func<ContentAssets, string>)(a => a.BluePng), L = 12.0, T = 45.0, W = 28.0, H = 28.0 },
                new { Tag = "green-l", Pick = (Func<ContentAssets, string>)(a => a.GreenPng), L = 45.0, T = 12.0, W = 32.0, H = 32.0 },
                new { Tag = "red-corner", Pick = (Func<ContentAssets, string>)(a => a.RedPng), L = 2.0, T = 2.0, W = 22.0, H = 22.0 },
                new { Tag = "blue-corner", Pick = (Func<ContentAssets, string>)(a => a.BluePng), L = 75.0, T = 2.0, W = 20.0, H = 20.0 },
                new { Tag = "green-corner", Pick = (Func<ContentAssets, string>)(a => a.GreenPng), L = 75.0, T = 75.0, W = 18.0, H = 18.0 }
            };

            for (int i = 0; i < specs.Length; i++)
            {
                var s = specs[i];
                Add(list, ref page, "pic-mx-src-" + s.Tag,
                    a => ContentHtml.PictureCreate(s.Pick(a), s.L, s.T, s.W, s.H),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "picture");
                        string g = GeoEq(n, s.L, s.T, s.W, s.H);
                        if (g != null) return g;
                        string err = ContentAssert.AttachPictureSrc(
                            ctx.Presentation, ctx.AfterCreate, ctx.ExportDir);
                        if (err != null) return err;
                        return string.IsNullOrEmpty(n.DataSrc) ? "无 data-src" : null;
                    });
            }
        }

        private static void AddPicMxReplaceKeepGeo(List<ContentCase> list, ref int page)
        {
            var pairs = new[]
            {
                new { Tag = "r2b", From = (Func<ContentAssets, string>)(a => a.RedPng), To = (Func<ContentAssets, string>)(a => a.BluePng) },
                new { Tag = "r2g", From = (Func<ContentAssets, string>)(a => a.RedPng), To = (Func<ContentAssets, string>)(a => a.GreenPng) },
                new { Tag = "b2r", From = (Func<ContentAssets, string>)(a => a.BluePng), To = (Func<ContentAssets, string>)(a => a.RedPng) },
                new { Tag = "b2g", From = (Func<ContentAssets, string>)(a => a.BluePng), To = (Func<ContentAssets, string>)(a => a.GreenPng) },
                new { Tag = "g2r", From = (Func<ContentAssets, string>)(a => a.GreenPng), To = (Func<ContentAssets, string>)(a => a.RedPng) },
                new { Tag = "g2b", From = (Func<ContentAssets, string>)(a => a.GreenPng), To = (Func<ContentAssets, string>)(a => a.BluePng) }
            };

            double[][] geos =
            {
                new[] { 10.0, 10.0, 30.0, 30.0 }, new[] { 40.0, 15.0, 35.0, 40.0 },
                new[] { 15.0, 45.0, 40.0, 25.0 }, new[] { 55.0, 20.0, 25.0, 35.0 }
            };

            for (int p = 0; p < pairs.Length; p++)
            {
                for (int g = 0; g < geos.Length; g++)
                {
                    var pair = pairs[p];
                    double l = geos[g][0], t = geos[g][1], w = geos[g][2], h = geos[g][3];
                    string name = "pic-mx-repl-keep-" + pair.Tag + "-" + g;
                    Add(list, ref page, name,
                        a => ContentHtml.PictureCreate(pair.From(a), l, t, w, h),
                        (a, ids) => ContentHtml.PictureUpdate(ids[0], pair.To(a)),
                        ctx => SrcChangedAndGeoKept(ctx));
                }
            }
        }

        private static void AddPicMxReplaceChangeGeo(List<ContentCase> list, ref int page)
        {
            var moves = new[]
            {
                new { Tag = "shrink", Fl = 20.0, Ft = 20.0, Fw = 40.0, Fh = 40.0, Tl = 25.0, Tt = 25.0, Tw = 25.0, Th = 25.0 },
                new { Tag = "grow", Fl = 30.0, Ft = 30.0, Fw = 20.0, Fh = 20.0, Tl = 15.0, Tt = 15.0, Tw = 45.0, Th = 40.0 },
                new { Tag = "move-r", Fl = 10.0, Ft = 20.0, Fw = 30.0, Fh = 30.0, Tl = 50.0, Tt = 20.0, Tw = 30.0, Th = 30.0 },
                new { Tag = "move-d", Fl = 20.0, Ft = 10.0, Fw = 35.0, Fh = 25.0, Tl = 20.0, Tt = 45.0, Tw = 35.0, Th = 25.0 },
                new { Tag = "wide", Fl = 25.0, Ft = 25.0, Fw = 25.0, Fh = 35.0, Tl = 10.0, Tt = 30.0, Tw = 55.0, Th = 20.0 },
                new { Tag = "tall", Fl = 20.0, Ft = 20.0, Fw = 40.0, Fh = 20.0, Tl = 30.0, Tt = 10.0, Tw = 20.0, Th = 55.0 },
                new { Tag = "corner", Fl = 40.0, Ft = 40.0, Fw = 30.0, Fh = 30.0, Tl = 5.0, Tt = 5.0, Tw = 25.0, Th = 25.0 },
                new { Tag = "center", Fl = 5.0, Ft = 5.0, Fw = 20.0, Fh = 20.0, Tl = 30.0, Tt = 30.0, Tw = 35.0, Th = 35.0 },
                new { Tag = "r2b-geo", Fl = 15.0, Ft = 35.0, Fw = 30.0, Fh = 30.0, Tl = 45.0, Tt = 15.0, Tw = 35.0, Th = 35.0 },
                new { Tag = "b2g-geo", Fl = 50.0, Ft = 10.0, Fw = 25.0, Fh = 40.0, Tl = 20.0, Tt = 40.0, Tw = 40.0, Th = 25.0 }
            };

            for (int i = 0; i < moves.Length; i++)
            {
                var m = moves[i];
                Func<ContentAssets, string> from = i % 2 == 0
                    ? (Func<ContentAssets, string>)(a => a.RedPng)
                    : (a => a.BluePng);
                Func<ContentAssets, string> to = i % 2 == 0
                    ? (Func<ContentAssets, string>)(a => a.GreenPng)
                    : (a => a.RedPng);
                Add(list, ref page, "pic-mx-repl-geo-" + m.Tag,
                    a => ContentHtml.PictureCreate(from(a), m.Fl, m.Ft, m.Fw, m.Fh),
                    (a, ids) => ContentHtml.PictureUpdate(ids[0], to(a), m.Tl, m.Tt, m.Tw, m.Th),
                    ctx =>
                    {
                        PptHtmlShapeNode after = PreferCreated(ctx, "picture", useReplace: true);
                        return GeoEq(after, m.Tl, m.Tt, m.Tw, m.Th);
                    });
            }
        }

        private static void AddPicMxRotations(List<ContentCase> list, ref int page)
        {
            double[] angles = { 0, 15, 30, 45, 90, 135, 180, 270, -15, -45 };
            double[][] geos =
            {
                new[] { 20.0, 20.0, 30.0, 30.0 },
                new[] { 40.0, 15.0, 25.0, 35.0 }
            };

            for (int ai = 0; ai < angles.Length; ai++)
            {
                for (int gi = 0; gi < geos.Length; gi++)
                {
                    double ang = angles[ai];
                    double l = geos[gi][0], t = geos[gi][1], w = geos[gi][2], h = geos[gi][3];
                    string angTag = ang < 0
                        ? "m" + ((int)(-ang)).ToString(CultureInfo.InvariantCulture)
                        : ((int)ang).ToString(CultureInfo.InvariantCulture);
                    string name = "pic-mx-rot-" + angTag + "-g" + gi;
                    Add(list, ref page, name,
                        a => ContentHtml.PictureCreate(
                            a.BluePng, l, t, w, h,
                            extraAttrs: "data-rotation=\""
                                + ang.ToString(CultureInfo.InvariantCulture) + "\""),
                        null,
                        ctx =>
                        {
                            PptHtmlShapeNode n = PreferCreated(ctx, "picture");
                            string g = GeoEq(n, l, t, w, h, eps: 1.2);
                            if (g != null) return g;
                            if (n == null) return "无 picture";
                            if (!n.Rotation.HasValue)
                            {
                                return "旋转未读回";
                            }

                            double live = n.Rotation.Value;
                            // Office 常把负角归一到 0～360
                            double expect = ang;
                            while (expect < 0) expect += 360;
                            while (expect >= 360) expect -= 360;
                            double liveN = live;
                            while (liveN < 0) liveN += 360;
                            while (liveN >= 360) liveN -= 360;
                            if (Math.Abs(liveN - expect) > 2.5
                                && Math.Abs(liveN - expect - 360) > 2.5
                                && Math.Abs(liveN - expect + 360) > 2.5)
                            {
                                return "旋转期望 " + expect + " 实际 " + live;
                            }

                            return null;
                        });
                }
            }
        }

        private static void AddPicMxNegatives(List<ContentCase> list, ref int page)
        {
            Add(list, ref page, "pic-mx-neg-http",
                a => ContentHtml.Section(
                    "  <img data-shape-type=\"picture\" data-src=\"http://example.com/x.png\" "
                    + "style=\"position:absolute;left:10%;top:10%;width:20%;height:20%\" />"),
                null,
                null,
                expectApplyErrorContains: "不存在");

            Add(list, ref page, "pic-mx-neg-https",
                a => ContentHtml.Section(
                    "  <img data-shape-type=\"picture\" data-src=\"https://example.com/y.png\" "
                    + "style=\"position:absolute;left:12%;top:12%;width:20%;height:20%\" />"),
                null,
                null,
                expectApplyErrorContains: "不存在");

            Add(list, ref page, "pic-mx-neg-data-uri",
                a => ContentHtml.Section(
                    "  <img data-shape-type=\"picture\" "
                    + "data-src=\"data:image/png;base64,AAAA\" "
                    + "style=\"position:absolute;left:14%;top:14%;width:20%;height:20%\" />"),
                null,
                null,
                expectApplyErrorContains: "data");

            Add(list, ref page, "pic-mx-neg-missing-file",
                a => ContentHtml.PictureCreate(
                    Path.Combine(Path.GetTempPath(), "easywrite-pic-mx-missing-" + Guid.NewGuid().ToString("N") + ".png"),
                    16, 16, 20, 20),
                null,
                null,
                expectApplyErrorContains: "不存在");

            Add(list, ref page, "pic-mx-neg-empty-src",
                a => ContentHtml.Section(
                    "  <img data-shape-type=\"picture\" data-src=\"\" "
                    + "style=\"position:absolute;left:18%;top:18%;width:20%;height:20%\" />"),
                null,
                null,
                expectApplyErrorContains: "data-src");
        }

        private static string SrcChangedAndGeoKept(ContentAssertContext ctx)
        {
            PptHtmlShapeNode before = PreferCreated(ctx, "picture", useReplace: false);
            PptHtmlShapeNode after = PreferCreated(ctx, "picture", useReplace: true);
            if (before == null || after == null) return "换图前后缺 picture";
            if (!ContentAssert.GeoClose(before.Style, after.Style, 1.0))
            {
                return "换图应保几何 before=" + before.Style + " after=" + after.Style;
            }

            string err = ContentAssert.AttachPictureSrc(
                ctx.Presentation, ctx.AfterReplace, Path.Combine(ctx.ExportDir, "after"));
            if (err != null) return "导出后: " + err;
            if (string.IsNullOrEmpty(before.DataSrc) || string.IsNullOrEmpty(after.DataSrc))
            {
                return "data-src 空";
            }

            return string.Equals(before.DataSrc, after.DataSrc, StringComparison.OrdinalIgnoreCase)
                ? "换图后 data-src 未变"
                : null;
        }

        private static string GeoEq(
            PptHtmlShapeNode n,
            double l,
            double t,
            double w,
            double h,
            double eps = 1.0)
        {
            if (n == null) return "无 picture";
            if (!ContentAssert.TryParseGeo(n.Style, out double al, out double at, out double aw, out double ah))
            {
                return "无几何: " + n.Style;
            }

            if (Math.Abs(al - l) > eps || Math.Abs(at - t) > eps
                || Math.Abs(aw - w) > eps || Math.Abs(ah - h) > eps)
            {
                return "几何偏离 期望 L" + l + " T" + t + " W" + w + " H" + h
                    + " 实际 " + n.Style;
            }

            return null;
        }

        private static string Tag(double v)
        {
            return ((int)Math.Round(v)).ToString(CultureInfo.InvariantCulture);
        }
    }
}
