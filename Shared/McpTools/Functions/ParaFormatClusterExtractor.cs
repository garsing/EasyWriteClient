using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WordAddIn1.DocumentMapping.CodeResolve;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// I15：源文档 eligible P_ 按 para_format 指纹聚类，无语义标签。
    /// </summary>
    public static class ParaFormatClusterExtractor
    {
        public static Dictionary<string, object> Extract(Word.Document document, bool disallowBackendApi = false)
        {
            ProcessDocumentOptions options = ProcessDocumentOptions.ForGetDocumentContent("source_para_clusters");
            options.DisallowBackendApi = disallowBackendApi;
            WordDocumentExtractor.ProcessDocument(document, options);

            var groups = new Dictionary<string, ClusterAcc>(StringComparer.Ordinal);
            IReadOnlyList<IReadOnlyList<string>> mapping = DocumentState.ParagraphNameMapping;
            if (mapping != null)
            {
                foreach (IReadOnlyList<string> row in mapping)
                {
                    if (row == null || row.Count == 0)
                    {
                        continue;
                    }

                    string code = row[0];
                    if (string.IsNullOrEmpty(code) || !code.StartsWith("P_", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (!ParagraphCodeResolver.TryResolveParagraphRange(
                            document,
                            code,
                            out Word.Range range,
                            out _,
                            out _))
                    {
                        continue;
                    }

                    if (!IsEligible(range))
                    {
                        continue;
                    }

                    Dictionary<string, object> paraFormat = ParaFormatReader.Extract(range);
                    string fingerprint = Fingerprint(paraFormat);
                    if (!groups.TryGetValue(fingerprint, out ClusterAcc acc))
                    {
                        acc = new ClusterAcc { ParaFormat = paraFormat };
                        groups[fingerprint] = acc;
                    }

                    acc.Count++;
                    string sample = SampleText(range);
                    if (!string.IsNullOrEmpty(sample) && acc.Samples.Count < 5)
                    {
                        acc.Samples.Add(sample);
                    }
                }
            }

            List<Dictionary<string, object>> clusters = groups.Values
                .OrderByDescending(x => x.Count)
                .Select((acc, i) =>
                {
                    var samples = acc.Samples.Take(5).ToList();
                    return new Dictionary<string, object>
                    {
                        ["cluster_id"] = "PC_" + (i + 1).ToString("00"),
                        ["para_format"] = acc.ParaFormat,
                        ["sample_texts"] = samples,
                        ["count"] = acc.Count
                    };
                })
                .ToList();

            return new Dictionary<string, object>
            {
                ["clusters"] = clusters,
                ["missing"] = false
            };
        }

        private static bool IsEligible(Word.Range range)
        {
            try
            {
                if (range.Tables == null || range.Tables.Count == 0)
                {
                    return true;
                }

                Word.Table table = range.Tables[1];
                return table.Rows.Count == 1 && table.Columns.Count == 1;
            }
            catch
            {
                return false;
            }
        }

        private static string SampleText(Word.Range range)
        {
            try
            {
                string text = (range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
                if (text.Length > 200)
                {
                    text = text.Substring(0, 200);
                }

                return text;
            }
            catch
            {
                return "";
            }
        }

        private static string Fingerprint(Dictionary<string, object> paraFormat)
        {
            if (paraFormat == null || paraFormat.Count == 0)
            {
                return "empty";
            }

            var keys = paraFormat.Keys.OrderBy(k => k, StringComparer.Ordinal).ToList();
            var sb = new StringBuilder();
            foreach (string key in keys)
            {
                sb.Append(key).Append('=').Append(paraFormat[key]).Append(';');
            }

            return sb.ToString();
        }

        private sealed class ClusterAcc
        {
            public Dictionary<string, object> ParaFormat;
            public int Count;
            public List<string> Samples = new List<string>();
        }
    }
}
