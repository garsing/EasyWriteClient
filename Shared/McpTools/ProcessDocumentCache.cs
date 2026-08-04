using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public sealed class BodyFingerprintDetail
    {
        public string Fingerprint { get; set; }
    }

    public sealed class ProcessDocumentCacheEntry
    {
        public string DocUuid { get; set; }
        public string BodyFingerprint { get; set; }
        public string ReadText { get; set; }
        public string Snapshot { get; set; }
        public string SegSnapshot { get; set; }
        public Dictionary<string, string> NameToContentMap { get; set; }
        public List<WordDocumentExtractor.ChunkInfo> ChunkInfo { get; set; }
        public Dictionary<int, string> TableIndexToIdMap { get; set; }
        public List<string> TableIdOrderSnapshot { get; set; }
        public Dictionary<int, string> ChartIndexToIdMap { get; set; }
        public List<string> ChartIdOrderSnapshot { get; set; }
        public List<List<string>> ChartIdMappingSnapshot { get; set; }
        public Dictionary<int, string> ImageIndexToIdMap { get; set; }
        public List<string> ImageIdOrderSnapshot { get; set; }
        public List<List<string>> ImageIdMappingSnapshot { get; set; }
        public string ProcessingMethod { get; set; }
        public DateTime CachedAtUtc { get; set; }
        public string LastSource { get; set; }
    }

    /// <summary>
    /// 对 WordOpenXML 中 w:body 做 stable 归一化后 SHA256，用于 ProcessDocument 缓存命中判定。
    /// </summary>
    public static class BodyXmlFingerprint
    {
        private static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static readonly HashSet<string> VolatileAttributeLocalNames =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "rsidR", "rsidRPr", "rsidP", "rsidDel", "rsidTr", "rsidSect",
                "rsidRDefault", "rsidPDefault", "rsidRPrDefault", "rsidRoot",
                "paraId", "textId", "editId", "anchor",
                "embed", "link", "relId", "tooltip", "editor", "modified", "spid"
            };

        private static readonly HashSet<string> VolatileElementLocalNames =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "proofErr", "lastRenderedPageBreak",
                "bookmarkStart", "bookmarkEnd",
                "commentRangeStart", "commentRangeEnd", "commentReference",
                "permStart", "permEnd",
                "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
                "customXmlInsRangeStart", "customXmlInsRangeEnd",
                "customXmlDelRangeStart", "customXmlDelRangeEnd",
                "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd",
                "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
                "fldChar", "smartTag", "smartTagPr", "dir", "bdo", "subDoc"
            };

        private static readonly HashSet<string> BlockRevisionRemovalLocalNames =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "del", "moveFrom"
            };

        private static readonly HashSet<string> BlockRevisionUnwrapLocalNames =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "ins", "moveTo"
            };

        private static readonly HashSet<string> FormattingPrLocalNames =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "pPr", "rPr", "tblPr", "trPr"
            };

        private static readonly HashSet<string> TcPrKeepChildLocalNames =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "gridSpan", "vMerge"
            };

        private static readonly HashSet<string> NormalizedPlaceholderLocalNames =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "drawing", "pict", "object", "control", "ruby", "rt"
            };

        public static BodyFingerprintDetail ComputeDetailed(Word.Document document)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            string docXml = document.Content.WordOpenXML;
            if (string.IsNullOrEmpty(docXml))
            {
                throw new InvalidOperationException("无法获取文档的 XML 内容");
            }

            XDocument xmlDoc = XDocument.Parse(docXml);
            XElement body = xmlDoc.Descendants(W + "body").FirstOrDefault();
            if (body == null)
            {
                throw new InvalidOperationException("无法找到文档主体 w:body");
            }

            XElement stableBody = BuildStableBody(body);
            string stableBodyXml = stableBody.ToString(SaveOptions.DisableFormatting);

            return new BodyFingerprintDetail
            {
                Fingerprint = HashUtf8(stableBodyXml)
            };
        }

        private static XElement BuildStableBody(XElement body)
        {
            var clone = new XElement(body);
            StripVolatileMetadata(clone);
            return clone;
        }

        private static void StripVolatileMetadata(XElement element)
        {
            foreach (XElement child in element.Elements().ToList())
            {
                StripVolatileMetadata(child);
            }

            foreach (XElement child in element.Elements().ToList())
            {
                string localName = child.Name.LocalName;

                if (BlockRevisionRemovalLocalNames.Contains(localName))
                {
                    child.Remove();
                    continue;
                }

                if (BlockRevisionUnwrapLocalNames.Contains(localName))
                {
                    UnwrapElement(child);
                    continue;
                }

                if (localName == "AlternateContent")
                {
                    UnwrapAlternateContent(child);
                    continue;
                }

                if (VolatileElementLocalNames.Contains(localName))
                {
                    child.Remove();
                    continue;
                }

                if (NormalizedPlaceholderLocalNames.Contains(localName))
                {
                    NormalizeToPlaceholder(child);
                }
            }

            foreach (XElement child in element.Elements().ToList())
            {
                string localName = child.Name.LocalName;
                if (FormattingPrLocalNames.Contains(localName))
                {
                    child.Remove();
                }
                else if (localName == "tcPr")
                {
                    TrimTcPr(child);
                }
            }

            StripVolatileAttributes(element);
        }

        private static void StripVolatileAttributes(XElement element)
        {
            var volatileAttributes = element.Attributes()
                .Where(IsVolatileAttribute)
                .ToList();
            foreach (XAttribute attr in volatileAttributes)
            {
                attr.Remove();
            }
        }

        private static bool IsVolatileAttribute(XAttribute attr)
        {
            if (attr == null)
            {
                return false;
            }

            string localName = attr.Name.LocalName;
            if (VolatileAttributeLocalNames.Contains(localName))
            {
                return true;
            }

            if (localName.StartsWith("rsid", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(localName, "id", StringComparison.OrdinalIgnoreCase))
            {
                string parentLocal = attr.Parent?.Name.LocalName;
                if (string.Equals(parentLocal, "hyperlink", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(parentLocal, "anchor", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(parentLocal, "subDoc", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static void UnwrapElement(XElement wrapper)
        {
            XElement parent = wrapper.Parent;
            if (parent == null)
            {
                wrapper.Remove();
                return;
            }

            var children = wrapper.Elements().ToList();
            foreach (XElement child in children)
            {
                child.Remove();
                parent.Add(child);
            }

            wrapper.Remove();
        }

        private static void UnwrapAlternateContent(XElement alternateContent)
        {
            XElement parent = alternateContent.Parent;
            if (parent == null)
            {
                alternateContent.Remove();
                return;
            }

            XElement choice = alternateContent.Elements()
                .FirstOrDefault(e => string.Equals(e.Name.LocalName, "Choice", StringComparison.OrdinalIgnoreCase))
                ?? alternateContent.Elements().FirstOrDefault();
            if (choice == null)
            {
                alternateContent.Remove();
                return;
            }

            var children = choice.Elements().ToList();
            foreach (XElement child in children)
            {
                child.Remove();
                parent.Add(child);
            }

            alternateContent.Remove();
        }

        private static void NormalizeToPlaceholder(XElement element)
        {
            element.RemoveAttributes();
            element.RemoveNodes();
        }

        private static void TrimTcPr(XElement tcPr)
        {
            foreach (XElement child in tcPr.Elements().ToList())
            {
                if (!TcPrKeepChildLocalNames.Contains(child.Name.LocalName))
                {
                    child.Remove();
                }
            }

            StripVolatileAttributes(tcPr);
        }

        private static string HashUtf8(string content)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(content ?? ""));
                return Convert.ToBase64String(hashBytes);
            }
        }
    }

    /// <summary>
    /// ProcessDocument 会话内缓存：body stable 指纹 HIT/MISS 判定与读写。
    /// </summary>
    public static class ProcessDocumentCache
    {
        private static int _loadedFingerprintVersion = -1;

        public static void EnsureFingerprintVersion()
        {
            if (_loadedFingerprintVersion != DocumentState.FingerprintVersion)
            {
                DocumentState.ClearProcessDocumentCache("fingerprint_version_changed");
                _loadedFingerprintVersion = DocumentState.FingerprintVersion;
            }
        }

        public static bool TryGetHit(
            string docUuid,
            BodyFingerprintDetail detail,
            ProcessDocumentOptions options,
            long bodyCheckMs,
            out WordDocumentExtractor.ProcessingResult result)
        {
            result = null;
            if (detail == null)
            {
                throw new ArgumentNullException(nameof(detail));
            }

            string source = options?.SnapshotSource ?? "(null)";
            string fingerprint = detail.Fingerprint;
            ProcessDocumentCacheEntry entry = DocumentState.LastProcessCache;

            if (entry == null || string.IsNullOrEmpty(entry.BodyFingerprint))
            {
                LogMiss(source, fingerprint, "cache_empty", bodyCheckMs);
                return false;
            }

            if (!string.Equals(entry.DocUuid, docUuid, StringComparison.Ordinal))
            {
                LogMiss(source, fingerprint, "cache_empty_or_doc_mismatch", bodyCheckMs);
                return false;
            }

            if (!string.Equals(entry.BodyFingerprint, fingerprint, StringComparison.Ordinal))
            {
                LogMiss(source, fingerprint, "body_changed", bodyCheckMs);
                return false;
            }

            if (options != null && options.BuildDisplayContent && !HasDisplayContent(entry))
            {
                LogMiss(source, fingerprint, "cache_miss_no_display", bodyCheckMs);
                return false;
            }

            result = BuildProcessingResultFromCache(entry);

            RestoreIdOrdersFromCache(entry);

            if (!string.Equals(entry.Snapshot, DocumentState.Snapshot, StringComparison.Ordinal)
                || !string.Equals(entry.SegSnapshot, DocumentState.SegSnapshot, StringComparison.Ordinal))
            {
                DocumentState.SetSnapshots(entry.Snapshot, entry.SegSnapshot, source);
            }

            LogHit(source, fingerprint, bodyCheckMs);
            return true;
        }

        public static void SaveAfterFullRead(
            string docUuid,
            BodyFingerprintDetail detail,
            WordDocumentExtractor.ProcessingResult result,
            string source)
        {
            if (result == null || detail == null || string.IsNullOrEmpty(detail.Fingerprint))
            {
                return;
            }

            var entry = new ProcessDocumentCacheEntry
            {
                DocUuid = docUuid,
                BodyFingerprint = detail.Fingerprint,
                ReadText = result.ReadText,
                Snapshot = result.Snapshot,
                SegSnapshot = result.SegSnapshot,
                NameToContentMap = CloneNameToContentMap(result.NameToContentMap),
                ChunkInfo = CloneChunkInfo(result.ChunkInfo),
                TableIndexToIdMap = CloneTableIndexMap(result.TableIndexToIdMap),
                TableIdOrderSnapshot = DocumentState.GetTableIdOrder()?.ToList(),
                ChartIndexToIdMap = CloneChartIndexMapFromOrder(DocumentState.GetChartIdOrder()),
                ChartIdOrderSnapshot = DocumentState.GetChartIdOrder()?.ToList(),
                ChartIdMappingSnapshot = DocumentState.GetChartIdMappingSnapshot(),
                ImageIndexToIdMap = CloneChartIndexMapFromOrder(DocumentState.GetImageIdOrder()),
                ImageIdOrderSnapshot = DocumentState.GetImageIdOrder()?.ToList(),
                ImageIdMappingSnapshot = DocumentState.GetImageIdMappingSnapshot(),
                ProcessingMethod = result.ProcessingMethod ?? "Local",
                CachedAtUtc = DateTime.UtcNow,
                LastSource = source ?? ""
            };

            DocumentState.SetLastProcessCache(entry);
        }

        private static bool HasDisplayContent(ProcessDocumentCacheEntry entry)
        {
            return entry.ChunkInfo != null
                && entry.ChunkInfo.Count > 0
                && !string.IsNullOrEmpty(entry.ChunkInfo[0].DisplayContent);
        }

        private static WordDocumentExtractor.ProcessingResult BuildProcessingResultFromCache(
            ProcessDocumentCacheEntry entry)
        {
            return new WordDocumentExtractor.ProcessingResult
            {
                ReadText = entry.ReadText,
                Snapshot = entry.Snapshot,
                SegSnapshot = entry.SegSnapshot,
                NameToContentMap = CloneNameToContentMap(entry.NameToContentMap),
                ChunkInfo = CloneChunkInfo(entry.ChunkInfo),
                TableIndexToIdMap = CloneTableIndexMap(entry.TableIndexToIdMap),
                ProcessingMethod = entry.ProcessingMethod ?? "Local"
            };
        }

        private static Dictionary<string, string> CloneNameToContentMap(Dictionary<string, string> src)
        {
            if (src == null)
            {
                return new Dictionary<string, string>();
            }

            return new Dictionary<string, string>(src);
        }

        private static Dictionary<int, string> CloneTableIndexMap(Dictionary<int, string> src)
        {
            if (src == null)
            {
                return new Dictionary<int, string>();
            }

            return new Dictionary<int, string>(src);
        }

        private static Dictionary<int, string> CloneChartIndexMapFromOrder(List<string> chartIdOrder)
        {
            var map = new Dictionary<int, string>();
            if (chartIdOrder == null)
            {
                return map;
            }

            for (int i = 0; i < chartIdOrder.Count; i++)
            {
                if (!string.IsNullOrEmpty(chartIdOrder[i]))
                {
                    map[i] = chartIdOrder[i];
                }
            }

            return map;
        }

        private static void RestoreIdOrdersFromCache(ProcessDocumentCacheEntry entry)
        {
            if (entry.TableIdOrderSnapshot != null)
            {
                DocumentState.SetTableIndexToIdMap(CloneChartIndexMapFromOrder(entry.TableIdOrderSnapshot));
            }

            if (entry.ChartIdOrderSnapshot != null)
            {
                DocumentState.SetChartIdOrder(entry.ChartIdOrderSnapshot);
            }

            if (entry.ChartIdMappingSnapshot != null)
            {
                DocumentState.SetChartIdMappingSnapshot(entry.ChartIdMappingSnapshot);
            }

            if (entry.ImageIdOrderSnapshot != null)
            {
                DocumentState.SetImageIdOrder(entry.ImageIdOrderSnapshot);
            }

            if (entry.ImageIdMappingSnapshot != null)
            {
                DocumentState.SetImageIdMappingSnapshot(entry.ImageIdMappingSnapshot);
            }
        }

        private static List<WordDocumentExtractor.ChunkInfo> CloneChunkInfo(
            List<WordDocumentExtractor.ChunkInfo> src)
        {
            if (src == null)
            {
                return new List<WordDocumentExtractor.ChunkInfo>();
            }

            return src.Select(c => new WordDocumentExtractor.ChunkInfo
            {
                ChunkIdx = c.ChunkIdx,
                Content = c.Content,
                DisplayContent = c.DisplayContent,
                ParagraphDisplayContent = c.ParagraphDisplayContent,
                Summary = c.Summary
            }).ToList();
        }

        public static void LogHit(string source, string fingerprint, long bodyCheckMs)
        {
            Debug.WriteLine(
                $"[ProcessCache] source={source} bodyFp={FpShort(fingerprint)} cache=HIT body_check_ms={bodyCheckMs}");
        }

        public static void LogMiss(string source, string fingerprint, string reason, long bodyCheckMs)
        {
            string msPart = bodyCheckMs > 0 ? $" body_check_ms={bodyCheckMs}" : "";
            Debug.WriteLine(
                $"[ProcessCache] source={source} bodyFp={FpShort(fingerprint)} cache=MISS reason={reason}{msPart}");
        }

        private static string FpShort(string fp)
        {
            if (string.IsNullOrEmpty(fp))
            {
                return "(null)";
            }

            return fp.Length <= 12 ? fp : fp.Substring(0, 12) + "...";
        }
    }
}
