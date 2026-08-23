using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// I15：源文档 eligible P_ 按 para_format 指纹聚类。读 OOXML，不逐段 COM Find。
    /// </summary>
    public static class ParaFormatClusterExtractor
    {
        private static readonly string[] FingerprintKeys =
        {
            "alignment",
            "first_line_indent",
            "left_indent",
            "right_indent",
            "hanging_indent",
            "line_spacing_rule",
            "line_spacing",
            "space_before",
            "space_after",
        };

        public static Dictionary<string, object> Extract(Word.Document document, bool disallowBackendApi = false)
        {
            ProcessDocumentOptions options = ProcessDocumentOptions.ForGetDocumentContent("source_para_clusters");
            options.DisallowBackendApi = disallowBackendApi;
            WordDocumentExtractor.ProcessDocument(document, options);

            var sw = Stopwatch.StartNew();
            OpenXmlPackage pkg = OpenXmlPackage.Load(document);
            XDocument xmlDoc = pkg.Xml;

            List<ParaFormatOoxmlReader.BodyParagraph> xmlParas =
                ParaFormatOoxmlReader.CollectBodyParagraphs(xmlDoc);
            Dictionary<string, ParaFormatOoxmlReader.StyleRec> styles =
                ParaFormatOoxmlReader.LoadStyles(xmlDoc);
            Dictionary<string, object> docDefaults = ParaFormatOoxmlReader.LoadDocDefaults(xmlDoc);

            List<string> codes = CollectParagraphCodes();
            Dictionary<string, ParaFormatOoxmlReader.BodyParagraph> aligned =
                AlignCodesToXml(codes, xmlParas);

            var groups = new Dictionary<string, ClusterAcc>(StringComparer.Ordinal);
            int used = 0;
            int skipped = 0;
            foreach (string code in codes)
            {
                if (!aligned.TryGetValue(code, out ParaFormatOoxmlReader.BodyParagraph xmlPara)
                    || xmlPara?.Paragraph == null)
                {
                    skipped++;
                    continue;
                }

                if (!xmlPara.Eligible)
                {
                    skipped++;
                    continue;
                }

                string stored = DocumentState.GetParagraphContent(code);
                string visible = ParaFormatOoxmlReader.NormalizeAlignText(stored);
                if (string.IsNullOrEmpty(visible))
                {
                    skipped++;
                    continue;
                }

                Dictionary<string, object> paraFormat = ParaFormatOoxmlReader.ExtractEffective(
                    xmlPara.Paragraph, styles, docDefaults);
                string fingerprint = Fingerprint(paraFormat);
                if (!groups.TryGetValue(fingerprint, out ClusterAcc acc))
                {
                    acc = new ClusterAcc { ParaFormat = paraFormat };
                    groups[fingerprint] = acc;
                }

                acc.Count++;
                used++;
                if (acc.Samples.Count < 5)
                {
                    string sample = visible.Length > 200 ? visible.Substring(0, 200) : visible;
                    if (!string.IsNullOrEmpty(sample) && !acc.Samples.Contains(sample))
                    {
                        acc.Samples.Add(sample);
                    }
                }
            }

            sw.Stop();
            System.Diagnostics.Debug.WriteLine(
                $"[ParaFormatCluster] ooxml P_={codes.Count} xml_p={xmlParas.Count} " +
                $"aligned={aligned.Count} used={used} skipped={skipped} elapsed_ms={sw.ElapsedMilliseconds}");

            return BuildPayload(groups);
        }

        private static Dictionary<string, object> BuildPayload(Dictionary<string, ClusterAcc> groups)
        {
            List<Dictionary<string, object>> clusters = groups.Values
                .OrderByDescending(x => x.Count)
                .ThenBy(x => Fingerprint(x.ParaFormat), StringComparer.Ordinal)
                .Select((acc, i) => new Dictionary<string, object>
                {
                    ["cluster_id"] = "PC_" + (i + 1).ToString("00"),
                    ["para_format"] = acc.ParaFormat,
                    ["sample_texts"] = acc.Samples.Take(5).ToList(),
                    ["count"] = acc.Count
                })
                .ToList();

            return new Dictionary<string, object>
            {
                ["clusters"] = clusters,
                ["missing"] = false
            };
        }

        private static List<string> CollectParagraphCodes()
        {
            var codes = new List<string>();
            IReadOnlyList<IReadOnlyList<string>> mapping = DocumentState.ParagraphNameMapping;
            if (mapping == null)
            {
                return codes;
            }

            foreach (IReadOnlyList<string> row in mapping)
            {
                if (row == null || row.Count == 0)
                {
                    continue;
                }

                string code = row[0];
                if (!string.IsNullOrEmpty(code) && code.StartsWith("P_", StringComparison.Ordinal))
                {
                    codes.Add(code);
                }
            }

            return codes;
        }

        private static Dictionary<string, ParaFormatOoxmlReader.BodyParagraph> AlignCodesToXml(
            List<string> codes,
            List<ParaFormatOoxmlReader.BodyParagraph> xmlParas)
        {
            var map = new Dictionary<string, ParaFormatOoxmlReader.BodyParagraph>(StringComparer.Ordinal);
            if (codes == null || xmlParas == null || xmlParas.Count == 0)
            {
                return map;
            }

            int xmlIdx = 0;
            foreach (string code in codes)
            {
                string target = ParaFormatOoxmlReader.NormalizeAlignText(
                    DocumentState.GetParagraphContent(code));
                ParaFormatOoxmlReader.BodyParagraph found = null;
                int windowEnd = Math.Min(xmlParas.Count, xmlIdx + 120);
                for (int i = xmlIdx; i < windowEnd; i++)
                {
                    ParaFormatOoxmlReader.BodyParagraph cand = xmlParas[i];
                    if (cand == null)
                    {
                        continue;
                    }

                    if (string.IsNullOrEmpty(target))
                    {
                        if (string.IsNullOrEmpty(cand.VisibleText))
                        {
                            found = cand;
                            xmlIdx = i + 1;
                            break;
                        }

                        continue;
                    }

                    if (string.Equals(cand.VisibleText, target, StringComparison.Ordinal))
                    {
                        found = cand;
                        xmlIdx = i + 1;
                        break;
                    }
                }

                if (found != null)
                {
                    map[code] = found;
                }
            }

            return map;
        }

        private static string Fingerprint(Dictionary<string, object> paraFormat)
        {
            if (paraFormat == null || paraFormat.Count == 0)
            {
                return "empty";
            }

            var sb = new StringBuilder();
            foreach (string key in FingerprintKeys)
            {
                if (!paraFormat.TryGetValue(key, out object val) || val == null)
                {
                    continue;
                }

                sb.Append(key).Append('=').Append(val).Append(';');
            }

            return sb.Length == 0 ? "empty" : sb.ToString();
        }

        private sealed class ClusterAcc
        {
            public Dictionary<string, object> ParaFormat;
            public int Count;
            public List<string> Samples = new List<string>();
        }
    }
}
