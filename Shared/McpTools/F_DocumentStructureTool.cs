using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>列出、自定义、套用样式。不改 display，不做目录。</summary>
    public static class F_DocumentStructureTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_document_structure"] = async (args) =>
            {
                try
                {
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    Word.Document doc = docHandle.Document;
                    string action = Arg(args, "action");
                    if (action == "list_spans" || action == "list_styles" || action == "customize" || action == "apply")
                    {
                        return await Task.FromResult(Run(doc, action, args, docHandle.DisallowBackendApi));
                    }

                    return new ToolResult { Success = false, Error = "action 须是 list_spans、list_styles、customize、apply。" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = ex.Message };
                }
            };
        }

        private static ToolResult Run(
            Word.Document doc,
            string action,
            Dictionary<string, object> args,
            bool disallowBackendApi)
        {
            OpenXmlPackage full = OpenXmlPackage.LoadFull(doc);
            OpenXmlPackage content = null;
            try
            {
                content = OpenXmlPackage.Load(doc);
            }
            catch (Exception)
            {
            }

            XElement stylesRoot = full.FindPartRoot("styles.xml", OpenXmlPackage.W + "styles");
            XElement body = content != null && content.Body != null ? content.Body : full.Body;
            List<StructureStyleRecord> styles = StructureStyleCatalog.Read(stylesRoot);

            if (action == "list_styles")
            {
                var rows = new List<Dictionary<string, object>>();
                foreach (StructureStyleRecord row in styles)
                {
                    rows.Add(row.ToCatalogRow());
                }

                return new ToolResult
                {
                    Success = true,
                    Data = new Dictionary<string, object>
                    {
                        ["styles"] = rows,
                        ["apply_rules"] = StructureStyleCatalog.ApplyRules
                    }
                };
            }

            if (action == "list_spans")
            {
                Dictionary<string, object> listed = StructureSpanLister.List(ReadDisplay(), body, styles);
                bool ok = listed.ContainsKey("ok") && listed["ok"] is bool bit && bit;
                return new ToolResult
                {
                    Success = ok,
                    Data = listed,
                    Error = ok ? null : (listed.ContainsKey("error") ? listed["error"]?.ToString() : "列出失败。")
                };
            }

            Word.Application app = doc.Application;
            bool screen = true;
            try
            {
                screen = app.ScreenUpdating;
                app.ScreenUpdating = false;
            }
            catch (Exception)
            {
                app = null;
            }

            try
            {
                if (action == "customize")
                {
                    Dictionary<string, object> edited = StructureStyleEditor.Customize(doc, args, styles);
                    bool ok = edited.ContainsKey("ok") && edited["ok"] is bool bit && bit;
                    return new ToolResult
                    {
                        Success = ok,
                        Data = edited,
                        Error = ok ? null : edited["error"]?.ToString()
                    };
                }

                var results = new List<Dictionary<string, object>>();
                int done = 0;
                foreach (Dictionary<string, object> item in Items(args))
                {
                    Dictionary<string, object> one = StructureStyleApplier.Apply(doc, item, styles, disallowBackendApi);
                    results.Add(one);
                    if (one.ContainsKey("ok") && one["ok"] is bool bit && bit)
                    {
                        done++;
                    }
                }

                return new ToolResult
                {
                    Success = done > 0,
                    Data = new Dictionary<string, object> { ["applied"] = results },
                    Error = done > 0 ? null : "没有套上。"
                };
            }
            finally
            {
                if (app != null)
                {
                    try
                    {
                        app.ScreenUpdating = screen;
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static string ReadDisplay()
        {
            ProcessDocumentCacheEntry cache = DocumentState.LastProcessCache;
            if (cache != null && cache.ChunkInfo != null && cache.ChunkInfo.Count > 0)
            {
                var sb = new StringBuilder();
                foreach (WordDocumentExtractor.ChunkInfo chunk in cache.ChunkInfo)
                {
                    if (chunk != null && !string.IsNullOrEmpty(chunk.DisplayContent))
                    {
                        sb.Append(chunk.DisplayContent);
                    }
                }

                if (sb.Length > 0)
                {
                    return sb.ToString();
                }
            }

            IReadOnlyList<string> parts = DocumentState.ChunkParagraphDisplayContents;
            if (parts == null || parts.Count == 0)
            {
                return "";
            }

            return string.Join("", parts);
        }

        private static IEnumerable<Dictionary<string, object>> Items(Dictionary<string, object> args)
        {
            if (args != null && args.TryGetValue("items", out object raw) && raw != null && !(raw is string))
            {
                if (raw is JArray array)
                {
                    foreach (JToken token in array)
                    {
                        yield return token.ToObject<Dictionary<string, object>>();
                    }

                    yield break;
                }

                if (raw is IEnumerable list)
                {
                    foreach (object entry in list)
                    {
                        Dictionary<string, object> item = AsDict(entry);
                        if (item != null)
                        {
                            yield return item;
                        }
                    }

                    yield break;
                }
            }

            yield return args ?? new Dictionary<string, object>();
        }

        private static Dictionary<string, object> AsDict(object raw)
        {
            if (raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            if (raw is JObject jo)
            {
                return jo.ToObject<Dictionary<string, object>>();
            }

            return null;
        }

        private static string Arg(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out object value) || value == null)
            {
                return "";
            }

            return value.ToString().Trim();
        }
    }
}
