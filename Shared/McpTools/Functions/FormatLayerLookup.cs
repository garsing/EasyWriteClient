using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Xml.Linq;
using WordAddIn1.DocumentMapping;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 段落最终层走读文映射上的 w:p，和 fs-p-ooxml 同一条。
    /// 读文没留下的段没有 P_。
    /// </summary>
    public static class FormatLayerLookup
    {
        public sealed class MappedParagraph
        {
            public string Code;
            public Dictionary<string, object> Direct;
            public Dictionary<string, object> Style;
            public Dictionary<string, object> Final;
        }

        private static List<WordReader.ParagraphMapSpan> _cachedMap;
        private static List<MappedParagraph> _cachedRows;
        private static string _cachedError;
        private static bool _hasCache;

        public static bool TryMap(
            Word.Document doc,
            bool disallowBackendApi,
            out List<MappedParagraph> rows,
            out string error)
        {
            rows = null;
            Stopwatch total = EasyWriteDiagnostics.IsEnabled(DebugCategory.Timing)
                ? Stopwatch.StartNew()
                : null;
            string detail = "fail=ensure";
            try
            {
                using (EasyWriteDiagnostics.Time("format_map.ensure"))
                {
                    error = Ensure(doc, disallowBackendApi);
                }

                if (error != null)
                {
                    return false;
                }

                if (_hasCache && ReferenceEquals(_cachedMap, WordReader.LastParagraphMap))
                {
                    rows = _cachedRows;
                    error = _cachedError;
                    detail = "cache=hit paragraphs=" + (rows == null ? 0 : rows.Count);
                    return rows != null;
                }

                List<WordDocumentExtractor.ParagraphMapProjection> projected;
                string display;
                List<DisplayParagraphIndexMap.Hit> hits;
                using (EasyWriteDiagnostics.Time("format_map.project"))
                {
                    projected = WordDocumentExtractor.ProjectParagraphMap(
                        WordReader.LastMappedReadText,
                        WordReader.LastParagraphMap);
                    display = DisplayParagraphIndexMap.CombinedParagraphDisplay();
                    hits = string.IsNullOrEmpty(display)
                        ? null
                        : DisplayParagraphIndexMap.Build(display);
                }

                if (string.IsNullOrEmpty(display))
                {
                    error = "没有段落 display。请先 F_get_document_content。";
                    detail = "fail=display";
                    return false;
                }

                if (projected.Count != hits.Count)
                {
                    error = "段落映射和 P_ 对不上：映射 " + projected.Count + "，P_ " + hits.Count;
                    Remember(error);
                    detail = "fail=align";
                    return false;
                }

                Dictionary<string, ParaFormatOoxmlReader.StyleRec> styles;
                Dictionary<string, object> defaults;
                using (EasyWriteDiagnostics.Time("format_map.styles"))
                {
                    if (!TryLoadStyles(doc, out styles, out defaults, out error))
                    {
                        detail = "fail=styles";
                        return false;
                    }
                }

                rows = new List<MappedParagraph>(projected.Count);
                using (EasyWriteDiagnostics.Time("format_map.layers"))
                {
                    for (int i = 0; i < projected.Count; i++)
                    {
                        XElement paragraph = FindParagraph(projected[i].SourceIndex);
                        if (paragraph == null)
                        {
                            error = "第 " + (i + 1) + " 段没有对应的 OOXML 段落";
                            rows = null;
                            Remember(error);
                            detail = "fail=layers";
                            return false;
                        }

                        ParaFormatOoxmlReader.ReadLayers(
                            paragraph,
                            styles,
                            defaults,
                            out Dictionary<string, object> direct,
                            out Dictionary<string, object> style,
                            out Dictionary<string, object> final);
                        rows.Add(new MappedParagraph
                        {
                            Code = hits[i].Code,
                            Direct = direct,
                            Style = style,
                            Final = final
                        });
                    }
                }

                _cachedRows = rows;
                _cachedError = null;
                _cachedMap = WordReader.LastParagraphMap;
                _hasCache = true;
                detail = "cache=miss paragraphs=" + rows.Count;
                return true;
            }
            finally
            {
                if (total != null)
                {
                    total.Stop();
                    EasyWriteDiagnostics.LogTiming("format_map.total", total.ElapsedMilliseconds, detail);
                }
            }
        }

        private static void Remember(string error)
        {
            _cachedRows = null;
            _cachedError = error;
            _cachedMap = WordReader.LastParagraphMap;
            _hasCache = true;
        }

        public static void ApplyParagraph(Word.Document doc, string code, Dictionary<string, object> entry)
        {
            if (doc == null || entry == null || string.IsNullOrEmpty(code))
            {
                return;
            }

            if (!TryMap(doc, false, out List<MappedParagraph> rows, out _))
            {
                return;
            }

            foreach (MappedParagraph row in rows)
            {
                if (!string.Equals(row.Code, code, StringComparison.Ordinal))
                {
                    continue;
                }

                entry["para_format"] = row.Final;
                entry["format_layers"] = new Dictionary<string, object>
                {
                    ["direct"] = row.Direct,
                    ["style"] = row.Style,
                    ["final"] = new Dictionary<string, object>(row.Final)
                };
                return;
            }
        }

        private static string Ensure(Word.Document doc, bool disallowBackendApi)
        {
            string prepared = InsertTargetLocatorHelper.TryProcessDocumentForLocate(doc, disallowBackendApi);
            if (prepared != null)
            {
                return prepared;
            }

            if (WordReader.LastParagraphMap != null)
            {
                return null;
            }

            try
            {
                ProcessDocumentOptions options = ProcessDocumentOptions.ForGetDocumentContent("format-map");
                options.ForceRefresh = true;
                options.DisallowBackendApi = disallowBackendApi;
                WordDocumentExtractor.ProcessDocument(doc, options);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            if (WordReader.LastParagraphMap == null)
            {
                return "没有段落映射";
            }

            return null;
        }

        private static XElement FindParagraph(int sourceIndex)
        {
            List<WordReader.ParagraphMapSpan> spans = WordReader.LastParagraphMap;
            if (spans == null || sourceIndex < 0)
            {
                return null;
            }

            for (int i = 0; i < spans.Count; i++)
            {
                WordReader.ParagraphMapSpan span = spans[i];
                if (span != null && span.Index == sourceIndex)
                {
                    return span.Paragraph;
                }
            }

            return null;
        }

        private static bool TryLoadStyles(
            Word.Document doc,
            out Dictionary<string, ParaFormatOoxmlReader.StyleRec> styles,
            out Dictionary<string, object> defaults,
            out string error)
        {
            styles = null;
            defaults = null;
            error = null;
            if (!ParaFormatOoxmlReader.TryReadOpenXml(doc, out XDocument xml, out error))
            {
                return false;
            }

            styles = ParaFormatOoxmlReader.LoadStyles(xml);
            defaults = ParaFormatOoxmlReader.LoadDocDefaults(xml);
            return true;
        }
    }
}
