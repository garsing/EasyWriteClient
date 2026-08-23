using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// I14：对源文档抽全表 catalog + standard（kb_table_id = 源 T_ stem）。
    /// </summary>
    public static class SourceTableFormatCatalogExtractor
    {
        public static void ExtractIntoSession(Word.Document document, DocumentSessionState session, bool disallowBackendApi = false)
        {
            if (session == null)
            {
                return;
            }

            var catalog = new List<Dictionary<string, object>>();
            var xmlById = new Dictionary<string, string>(StringComparer.Ordinal);
            ProcessDocumentOptions options = ProcessDocumentOptions.ForGetDocumentContent("source_table_catalog");
            options.DisallowBackendApi = disallowBackendApi;
            WordDocumentExtractor.ProcessDocument(document, options);

            List<string> tableIds = DocumentState.GetTableIdOrder();
            if (tableIds == null || tableIds.Count == 0)
            {
                session.TableFormatCatalog = catalog;
                session.TableFormatStandard = null;
                return;
            }

            int wordTableCount = 0;
            try
            {
                wordTableCount = document.Tables.Count;
            }
            catch
            {
                wordTableCount = 0;
            }

            for (int i = 0; i < tableIds.Count; i++)
            {
                string tableId = tableIds[i];
                if (string.IsNullOrEmpty(tableId))
                {
                    continue;
                }

                int comIndex = i + 1;
                if (comIndex > wordTableCount)
                {
                    break;
                }

                try
                {
                    Word.Table table = document.Tables[comIndex];
                    TableExtractDto dto = TableFormatExtractCore.Extract(table, document);
                    string xml = TableFormatXmlBuilder.BuildGeneralOnly(dto);
                    xmlById[tableId] = xml;

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
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        "[SourceTableFormatCatalogExtractor] " + tableId + ": " + ex.Message);
                }
            }

            session.TableFormatCatalog = catalog;
            session.TableFormatStandard = BuildStandard(catalog, xmlById);
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
            List<Dictionary<string, object>> catalog,
            Dictionary<string, string> xmlById)
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
