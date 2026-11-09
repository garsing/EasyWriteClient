using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;

namespace ExportImageCaptionTest
{
    internal static class Program
    {
        private static int _failed;

        private static int Main()
        {
            Console.OutputEncoding = Encoding.UTF8;
            string root = Path.Combine(Path.GetTempPath(), "ew-export-caption-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                ThreeImages(root);
                TwentyFive(root);
                OneHundredOne(root);
                SecondBatchFails(root);
                UnclearTruncateMissing(root);
                PlaceholderAndBareTag(root);
                ExcelListDoesNotCall(root);
            }
            finally
            {
                WordAddIn1.ExportImageCaption.RequestTextOverride = null;
                try
                {
                    Directory.Delete(root, true);
                }
                catch (Exception)
                {
                }
            }

            if (_failed == 0)
            {
                Console.WriteLine("过");
                return 0;
            }

            Console.WriteLine("败 " + _failed);
            return 1;
        }

        private static void ThreeImages(string root)
        {
            string dir = Path.Combine(root, "three");
            string html = Html(dir, 3, out string firstPath);
            long before = new FileInfo(firstPath).Length;
            var calls = new List<int>();
            WordAddIn1.ExportImageCaption.RequestTextOverride = paths =>
            {
                calls.Add(paths.Count);
                return Lines(paths.Count, i => "图" + i);
            };
            string updated = WordAddIn1.ExportImageCaption.ApplyHtml(Wrap(html));
            Check(calls.Count == 1 && calls[0] == 3, "3 张只请求一次");
            Check(updated.Contains("desc=\"图1\"") && updated.Contains("desc=\"图3\""), "3 张写回 desc");
            Check(new FileInfo(firstPath).Length == before, "原图字节不变");
        }

        private static void TwentyFive(string root)
        {
            string html = Html(Path.Combine(root, "twentyfive"), 25, out _);
            var calls = new List<int>();
            WordAddIn1.ExportImageCaption.RequestTextOverride = paths =>
            {
                calls.Add(paths.Count);
                return Lines(paths.Count, i => "批" + calls.Count + "-" + i);
            };
            string updated = WordAddIn1.ExportImageCaption.ApplyHtml(Wrap(html));
            Check(calls.Count == 2 && calls[0] == 20 && calls[1] == 5, "25 张分 20+5");
            Check(updated.Contains("desc=\"批1-1\"") && updated.Contains("desc=\"批2-5\""), "两批按总顺序写回");
            Check(updated.IndexOf("desc=\"批1-20\"", StringComparison.Ordinal) < updated.IndexOf("desc=\"批2-1\"", StringComparison.Ordinal), "第二批在第一批之后");
        }

        private static void OneHundredOne(string root)
        {
            string html = Html(Path.Combine(root, "hundred"), 101, out _);
            var calls = new List<int>();
            WordAddIn1.ExportImageCaption.RequestTextOverride = paths =>
            {
                calls.Add(paths.Count);
                return Lines(paths.Count, i => "有");
            };
            string updated = WordAddIn1.ExportImageCaption.ApplyHtml(Wrap(html));
            int sum = 0;
            foreach (int n in calls)
            {
                sum += n;
            }

            Check(calls.Count == 5 && sum == 100, "100 张 5 次，没有第 6 次");
            Check(CountOf(updated, "desc=\"有\"") == 100, "前 100 张有内容");
            Check(updated.Contains("desc=\"图片太多，没有打标内容\""), "第 101 张固定说明");
            Check(CountOf(updated, "desc=\"图片太多，没有打标内容\"") == 1, "只有超出的那张写固定说明");
        }

        private static void SecondBatchFails(string root)
        {
            string html = Html(Path.Combine(root, "fail"), 25, out string first);
            long before = new FileInfo(first).Length;
            int n = 0;
            WordAddIn1.ExportImageCaption.RequestTextOverride = paths =>
            {
                n++;
                if (n == 2)
                {
                    throw new InvalidOperationException("失败");
                }

                return Lines(paths.Count, i => "不该留下");
            };
            var data = new Dictionary<string, object>
            {
                ["message"] = "抽取成功",
                ["document_content"] = html
            };
            WordAddIn1.ExportImageCaption.ApplyDocumentContent(data);
            string updated = (string)data["document_content"];
            Check(!updated.Contains("desc="), "第二次失败则模型说明都不写");
            Check((string)data["message"] == "抽取成功", "导出仍成功");
            Check(new FileInfo(first).Length == before, "失败后文件还在");

            string more = Html(Path.Combine(root, "fail101"), 101, out _);
            n = 0;
            WordAddIn1.ExportImageCaption.RequestTextOverride = paths =>
            {
                n++;
                if (n == 2)
                {
                    throw new InvalidOperationException("失败");
                }

                return Lines(paths.Count, i => "不该留下");
            };
            string updatedMore = WordAddIn1.ExportImageCaption.ApplyHtml(Wrap(more));
            Check(!updatedMore.Contains("desc=\"不该留下\""), "失败时前 100 张没有模型说明");
            Check(updatedMore.Contains("desc=\"图片太多，没有打标内容\""), "失败时第 101 张仍是固定说明");
        }

        private static void UnclearTruncateMissing(string root)
        {
            string html = Html(Path.Combine(root, "parse"), 3, out _);
            WordAddIn1.ExportImageCaption.RequestTextOverride = paths =>
                "1. 看不清\n2. " + new string('字', 25) + "\n";
            string updated = WordAddIn1.ExportImageCaption.ApplyHtml(Wrap(html));
            Check(updated.Contains("desc=\"看不清\""), "看不清原样写入");
            Check(updated.Contains("desc=\"" + new string('字', 20) + "\""), "超过 20 个字符只留前 20 个");
            Check(CountOf(updated, "desc=") == 2, "缺一行的那张不写 desc");
        }

        private static void PlaceholderAndBareTag(string root)
        {
            string dir = Path.Combine(root, "skip");
            Directory.CreateDirectory(dir);
            string real = Path.Combine(dir, "real.png");
            SavePng(real);
            string html =
                "前文<image id=\"I_x\" src=\"lightweight/placeholder_1.png\"/>" +
                "<image slide_id=\"1\" shape_id=\"s1\"/>" +
                "<image id=\"I_y\" src=\"" + real.Replace('\\', '/') + "\"/>后文";
            var calls = new List<int>();
            WordAddIn1.ExportImageCaption.RequestTextOverride = paths =>
            {
                calls.Add(paths.Count);
                return "1. 只有这张";
            };
            string updated = WordAddIn1.ExportImageCaption.ApplyHtml(Wrap(html));
            Check(calls.Count == 1 && calls[0] == 1, "占位和无 src 不送");
            Check(updated.Contains("src=\"lightweight/placeholder_1.png\"/>"), "占位标签不变");
            Check(updated.Contains("shape_id=\"s1\"/>"), "无 src 的标签不变");
            Check(updated.Contains("desc=\"只有这张\""), "真实文件写上 desc");
        }

        private static void ExcelListDoesNotCall(string root)
        {
            int calls = 0;
            WordAddIn1.ExportImageCaption.RequestTextOverride = paths =>
            {
                calls++;
                return "1. 不该";
            };
            var listData = new Dictionary<string, object>
            {
                ["action"] = "list",
                ["message"] = "抽取成功",
                ["images"] = new List<Dictionary<string, object>>()
            };
            WordAddIn1.ExportImageCaption.ApplyExcelExport(listData);
            Check(calls == 0, "list 不调用接口");

            string path = Path.Combine(root, "excel.png");
            SavePng(path);
            var row = new Dictionary<string, object>
            {
                ["image_id"] = "I_a1b2c3d4",
                ["path"] = path.Replace('\\', '/')
            };
            var exportData = new Dictionary<string, object>
            {
                ["action"] = "export",
                ["message"] = "抽取成功",
                ["images"] = new List<Dictionary<string, object>> { row }
            };
            WordAddIn1.ExportImageCaption.RequestTextOverride = paths => "1. 季度营收柱状图";
            WordAddIn1.ExportImageCaption.ApplyExcelExport(exportData);
            Check((string)row["desc"] == "季度营收柱状图", "Excel 字段 desc");
        }

        private static string Wrap(string html)
        {
            return html;
        }

        private static string Html(string dir, int count, out string firstPath)
        {
            Directory.CreateDirectory(dir);
            firstPath = Path.Combine(dir, "1.png");
            var builder = new StringBuilder();
            for (int i = 1; i <= count; i++)
            {
                string path = Path.Combine(dir, i + ".png");
                SavePng(path);
                builder.Append("<image id=\"I_").Append(i).Append("\" src=\"").Append(path.Replace('\\', '/')).Append("\"/>");
            }

            return builder.ToString();
        }

        private static void SavePng(string path)
        {
            using (var bitmap = new Bitmap(8, 6))
            {
                bitmap.SetPixel(0, 0, Color.Blue);
                bitmap.Save(path, ImageFormat.Png);
            }
        }

        private static string Lines(int count, Func<int, string> text)
        {
            var builder = new StringBuilder();
            for (int i = 1; i <= count; i++)
            {
                builder.Append(i).Append(". ").Append(text(i)).Append('\n');
            }

            return builder.ToString();
        }

        private static int CountOf(string text, string piece)
        {
            int count = 0;
            int index = 0;
            while (index >= 0 && index < text.Length)
            {
                index = text.IndexOf(piece, index, StringComparison.Ordinal);
                if (index < 0)
                {
                    break;
                }

                count++;
                index += piece.Length;
            }

            return count;
        }

        private static void Check(bool ok, string name)
        {
            if (ok)
            {
                Console.WriteLine("过 " + name);
                return;
            }

            _failed++;
            Console.WriteLine("败 " + name);
        }
    }
}
