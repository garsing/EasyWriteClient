using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// I14：对源文档抽全表 catalog + standard。文档级一次 WordOpenXML，失败直接抛。
    /// </summary>
    public static class SourceTableFormatCatalogExtractor
    {
        private static readonly XNamespace W = OpenXmlPackage.W;

        public static void ExtractIntoSession(Word.Document document, DocumentSessionState session, bool disallowBackendApi = false)
        {
            if (session == null)
            {
                return;
            }

            ProcessDocumentOptions options = ProcessDocumentOptions.ForGetDocumentContent("source_table_catalog");
            options.DisallowBackendApi = disallowBackendApi;
            WordDocumentExtractor.ProcessDocument(document, options);

            List<string> tableIds = DocumentState.GetTableIdOrder();
            if (tableIds == null || tableIds.Count == 0)
            {
                session.TableFormatCatalog = new List<Dictionary<string, object>>();
                session.TableFormatStandard = null;
                return;
            }

            OpenXmlPackage pkg = OpenXmlPackage.Load(document);
            List<XElement> tables = pkg.Body.Descendants(W + "tbl").ToList();
            if (tables.Count != tableIds.Count)
            {
                throw new InvalidOperationException(
                    "WordOpenXML 表数(" + tables.Count + ") 与 T_ 数(" + tableIds.Count + ") 不一致");
            }

            var catalog = new List<Dictionary<string, object>>();
            for (int i = 0; i < tableIds.Count; i++)
            {
                string tableId = tableIds[i];
                if (string.IsNullOrEmpty(tableId))
                {
                    continue;
                }

                TableExtractDto dto = TableFormatExtractCore.Extract(tables[i]);
                string xml = TableFormatXmlBuilder.BuildGeneralOnly(dto);

                var entry = new Dictionary<string, object>
                {
                    ["kb_table_id"] = tableId,
                    ["rows"] = dto.Rows,
                    ["cols"] = dto.Cols,
                    ["xml_content"] = xml
                };
                if (!string.IsNullOrEmpty(dto.Style))
                {
                    entry["Style"] = dto.Style;
                }

                if (dto.FontStats != null)
                {
                    if (!string.IsNullOrEmpty(dto.FontStats.TableFontName))
                    {
                        entry["TableFontName"] = dto.FontStats.TableFontName;
                    }

                    if (dto.FontStats.TableFontSize > 0)
                    {
                        entry["TableFontSize"] = dto.FontStats.TableFontSize;
                    }

                    if (!string.IsNullOrEmpty(dto.FontStats.TableFontColor))
                    {
                        entry["TableFontColor"] = dto.FontStats.TableFontColor;
                    }
                }

                catalog.Add(entry);
            }

            session.TableFormatCatalog = catalog;
            session.TableFormatStandard = BuildStandard(catalog);
        }

        public static string FindXml(DocumentSessionState session, string kbTableId)
        {
            if (session?.TableFormatCatalog == null || string.IsNullOrEmpty(kbTableId))
            {
                return null;
            }

            foreach (Dictionary<string, object> row in session.TableFormatCatalog)
            {
                if (row == null)
                {
                    continue;
                }

                string id = row.TryGetValue("kb_table_id", out object raw) ? raw?.ToString() : null;
                if (string.Equals(id, kbTableId, StringComparison.Ordinal)
                    && row.TryGetValue("xml_content", out object xmlRaw)
                    && xmlRaw != null)
                {
                    return xmlRaw.ToString();
                }
            }

            return null;
        }

        private static Dictionary<string, object> BuildStandard(
            List<Dictionary<string, object>> catalog)
        {
            if (catalog == null || catalog.Count == 0)
            {
                return null;
            }

            Dictionary<string, object> first = catalog[0];
            var standard = new Dictionary<string, object>(StringComparer.Ordinal);
            CopyIfPresent(first, standard, "TableFontName");
            CopyIfPresent(first, standard, "TableFontSize");
            CopyIfPresent(first, standard, "TableFontColor");
            CopyIfPresent(first, standard, "Style");
            if (first.TryGetValue("xml_content", out object xml) && xml != null)
            {
                standard["xml_content"] = xml.ToString();
            }

            return standard;
        }

        private static void CopyIfPresent(
            Dictionary<string, object> src,
            Dictionary<string, object> dest,
            string key)
        {
            if (src.TryGetValue(key, out object raw) && raw != null)
            {
                dest[key] = raw;
            }
        }
    }
}
