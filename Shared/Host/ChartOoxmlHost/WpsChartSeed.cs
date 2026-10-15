using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using WordAddIn1.OpenFiles;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.ChartOoxmlHost
{
    /// <summary>WPS 建图：最小 OOXML 种子 InsertFile / 拷图，再改包灌数。不在 WPS 上 AddChart。</summary>
    internal static class WpsChartSeed
    {
        public static bool TryCreate(object document, ChartConfig config, out string chartId, out string error)
        {
            chartId = null;
            error = null;
            if (document == null || config == null || config.Chart == null)
            {
                error = "建图配置空";
                return false;
            }

            if (!ChartConfigSyntax.TryCanonicalType(config.Chart.Type, out string type, out error))
            {
                return false;
            }

            string seedPath = WordChartPackageHost.NewTempDocx("seed-" + type);
            try
            {
                WriteSeedDocx(seedPath, type);
                if (!TryInsertSeed(document, seedPath, out error))
                {
                    return false;
                }

                List<object> shapes = WordChartPackageHost.CollectChartShapes(document);
                if (shapes.Count == 0)
                {
                    error = "种子插入后没有图";
                    return false;
                }

                int ordinal = shapes.Count - 1;
                object shape = shapes[ordinal];
                var occupied = new HashSet<string>(StringComparer.Ordinal);
                foreach (object existing in shapes)
                {
                    if (WordChartPackageHost.TryReadShapeTitle(existing, out string title)
                        && ChartTitleStampHelper.TryParse(title, out string existingId, out _))
                    {
                        occupied.Add(existingId);
                    }
                }

                chartId = ChartTitleStampHelper.NewRandomChartId(occupied);
                if (string.IsNullOrEmpty(chartId)
                    || !ChartTitleStampHelper.TryWriteStamp(
                        new ChartResolveHelper.BodyChart { Raw = shape },
                        chartId))
                {
                    error = "图表编号写入形状标题失败";
                    try
                    {
                        WordChartPackageHost.TryDeleteChartOnDoc(document, ordinal, out _);
                    }
                    catch
                    {
                    }

                    chartId = null;
                    return false;
                }

                ChartPartXml.SeriesGrid grid = ChartOoxmlFacade.GridFromChartInfo(config.Chart);
                grid.Type = type;
                if (!ChartOoxmlFacade.TryRewriteDataAndCopyBack(document, ordinal, grid, pinColors: false, out error))
                {
                    return false;
                }

                return true;
            }
            finally
            {
                WordChartPackageHost.TryDeleteFile(seedPath);
            }
        }

        private static bool TryInsertSeed(object document, string seedPath, out string error)
        {
            error = null;
            try
            {
                Word.Document word = document as Word.Document;
                if (word != null)
                {
                    Word.Range end = word.Content;
                    end.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    end.InsertFile(seedPath);
                    return true;
                }

                object content = WpsCom.GetProperty(document, "Content");
                WpsCom.Invoke(content, "Collapse", 0);
                WpsCom.Invoke(content, "InsertFile", seedPath);
                return true;
            }
            catch (Exception ex)
            {
                error = "插入种子失败: " + ex.Message;
                return false;
            }
        }

        private static void WriteSeedDocx(string path, string type)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            using (FileStream fs = File.Create(path))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                WriteEntry(zip, "[Content_Types].xml", ContentTypesXml());
                WriteEntry(zip, "_rels/.rels", RootRelsXml());
                WriteEntry(zip, "word/document.xml", DocumentXml());
                WriteEntry(zip, "word/_rels/document.xml.rels", DocumentRelsXml());
                WriteEntry(zip, "word/charts/chart1.xml", ChartXml(type));
            }
        }

        private static void WriteEntry(ZipArchive zip, string name, string xml)
        {
            ZipArchiveEntry e = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (var w = new StreamWriter(e.Open(), new UTF8Encoding(false)))
            {
                w.Write(xml);
            }
        }

        private static string ContentTypesXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "<Override PartName=\"/word/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>"
                + "</Types>";
        }

        private static string RootRelsXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>";
        }

        private static string DocumentRelsXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"charts/chart1.xml\"/>"
                + "</Relationships>";
        }

        private static string DocumentXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
                + " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\""
                + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
                + " xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
                + "<w:body><w:p><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">"
                + "<wp:extent cx=\"3657600\" cy=\"2743200\"/>"
                + "<wp:docPr id=\"1\" name=\"Chart 1\"/>"
                + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
                + "<c:chart r:id=\"rId1\"/>"
                + "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
                + "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr>"
                + "</w:body></w:document>";
        }

        private static string ChartXml(string type)
        {
            string plot;
            if (type == "bar")
            {
                plot = BarPlot("bar");
            }
            else if (type == ChartConfigSyntax.TypeLine)
            {
                plot = LinePlot();
            }
            else if (type == "pie2d")
            {
                plot = PiePlot(false);
            }
            else if (type == ChartConfigSyntax.TypePie3d)
            {
                plot = PiePlot(true);
            }
            else
            {
                plot = BarPlot("col");
            }

            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\""
                + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
                + "<c:chart><c:plotArea>" + plot + "</c:plotArea></c:chart></c:chartSpace>";
        }

        private static string Ser()
        {
            return "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
                + "<c:tx><c:v>系列1</c:v></c:tx>"
                + "<c:cat><c:strLit><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>类别1</c:v></c:pt></c:strLit></c:cat>"
                + "<c:val><c:numLit><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>1</c:v></c:pt></c:numLit></c:val>"
                + "</c:ser>";
        }

        private static string BarPlot(string dir)
        {
            return "<c:barChart><c:barDir val=\"" + dir + "\"/><c:grouping val=\"clustered\"/>"
                + Ser()
                + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>"
                + CatValAxes();
        }

        private static string LinePlot()
        {
            return "<c:lineChart><c:grouping val=\"standard\"/>" + Ser()
                + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>"
                + CatValAxes();
        }

        private static string PiePlot(bool d3)
        {
            string name = d3 ? "pie3DChart" : "pieChart";
            return "<c:" + name + ">" + Ser() + "</c:" + name + ">";
        }

        private static string CatValAxes()
        {
            return "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>"
                + "<c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>"
                + "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>"
                + "<c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>";
        }
    }
}
