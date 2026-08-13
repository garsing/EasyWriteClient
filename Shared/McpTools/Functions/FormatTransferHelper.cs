using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Xml.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 格式转移：读入与格式提取后拉取远程 subtype_format_mapping 与 section_page_layout；
    /// 先应用页布局 mode，再对整篇应用 body 字符格式（含表格单元格文字），再按 detailed_subtype 对齐标题。
    /// 表格 standard 粗迁代码保留但默认不调用（正文 body 已覆盖表内字体；见 ApplyTableFormatStandardInCoarseTransfer）。
    /// 卡点诊断：config App.DebugCategories 含 format_transfer 时输出逐步 enter/exit + elapsed_ms。
    /// </summary>
    public static class FormatTransferHelper
    {
        /// <summary>
        /// false：粗迁不拉取/不 apply table_format_standard（表内字体随 body 整篇 apply）。
        /// 相关方法保留供后续精迁或单独启用。
        /// </summary>
        private const bool ApplyTableFormatStandardInCoarseTransfer = false;

        private static bool FtDebugEnabled => EasyWriteDiagnostics.IsEnabled(DebugCategory.FormatTransfer);

        private static void FtLog(string message)
        {
            EasyWriteDiagnostics.Log(DebugCategory.FormatTransfer, message);
        }

        /// <summary>粗迁步骤作用域：enter 立即打点，dispose 时打 exit + elapsed_ms（便于定位卡住的最后一步）。</summary>
        private sealed class FtStep : IDisposable
        {
            private readonly string _step;
            private readonly Stopwatch _sw;
            private bool _disposed;

            public FtStep(string step, string detail = null)
            {
                _step = step ?? "";
                if (!FtDebugEnabled)
                {
                    _sw = null;
                    return;
                }

                _sw = Stopwatch.StartNew();
                if (string.IsNullOrEmpty(detail))
                {
                    FtLog($"ENTER step={_step}");
                }
                else
                {
                    FtLog($"ENTER step={_step} {detail}");
                }
            }

            public void Dispose()
            {
                if (_disposed || _sw == null)
                {
                    return;
                }

                _disposed = true;
                _sw.Stop();
                FtLog($"EXIT  step={_step} elapsed_ms={_sw.ElapsedMilliseconds}");
            }
        }

        /// <summary>
        /// 执行格式转移全流程（仅当前活动文档，D11）。
        /// </summary>
        public static async Task<ToolResult> RunFormatTransferAsync(
            Word.Application wordApp,
            string targetDocumentName,
            string targetKnowledgeBaseUuid = null,
            string targetStorageDocUuid = null,
            string pageLayoutContentMode = "auto",
            bool applyPageSetup = true,
            Word.Document document = null)
        {
            var totalSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
            FtLog(
                $"BEGIN RunFormatTransferAsync storage={targetStorageDocUuid ?? ""} " +
                $"doc={targetDocumentName ?? ""} kb={targetKnowledgeBaseUuid ?? ""} " +
                $"page_layout_content_mode={pageLayoutContentMode ?? "auto"} apply_page_setup={applyPageSetup}");

            var storageUuid = (targetStorageDocUuid ?? "").Trim();
            var docName = (targetDocumentName ?? "").Trim();
            var kbUuid = (targetKnowledgeBaseUuid ?? "").Trim();

            if (string.IsNullOrEmpty(storageUuid) && string.IsNullOrEmpty(docName))
            {
                FtLog("ABORT missing target identity");
                return new ToolResult { Success = false, Error = "缺少 target_storage_doc_uuid 或 target_document_name" };
            }

            if (string.IsNullOrEmpty(storageUuid) && string.IsNullOrEmpty(kbUuid))
            {
                FtLog("ABORT missing kb uuid for document_name");
                return new ToolResult { Success = false, Error = "按 document_name 定位时必须提供 target_knowledge_base_uuid" };
            }

            var userService = UserService.Instance;
            if (!userService.CheckLoginStatus())
            {
                FtLog("ABORT not logged in");
                return new ToolResult { Success = false, Error = "用户未登录，无法调用知识库接口" };
            }

            Word.Document doc = document;
            try
            {
                if (doc == null)
                {
                    if (wordApp == null)
                    {
                        FtLog("ABORT wordApp=null");
                        return new ToolResult { Success = false, Error = "Word 应用程序不可用" };
                    }

                    using (new FtStep("ActiveDocument"))
                    {
                        try
                        {
                            doc = wordApp.ActiveDocument;
                        }
                        catch (Exception ex)
                        {
                            FtLog($"ActiveDocument exception: {ex.Message}");
                            doc = null;
                        }
                    }
                }

                if (doc == null)
                {
                    FtLog("ABORT no active document");
                    return new ToolResult { Success = false, Error = "无活动文档" };
                }

                if (wordApp == null)
                {
                    try
                    {
                        wordApp = doc.Application;
                    }
                    catch (Exception ex)
                    {
                        FtLog($"doc.Application exception: {ex.Message}");
                        wordApp = null;
                    }
                }

                string docNameForLog = "";
                try
                {
                    docNameForLog = doc.Name ?? "";
                }
                catch
                {
                    // ignore COM
                }

                FtLog($"doc_ready name={docNameForLog}");

                using (new FtStep("ProcessDocument.read_and_format_extract"))
                {
                    RunReadAndFormatExtractionForTransfer(doc);
                }

                List<Dictionary<string, object>> targetMapping;
                using (new FtStep("FetchSubtypeFormatMapping"))
                {
                    targetMapping = await FetchSubtypeFormatMappingAsync(
                        docName, kbUuid, storageUuid).ConfigureAwait(false);
                }

                if (targetMapping == null)
                {
                    FtLog("ABORT FetchSubtypeFormatMapping returned null");
                    return new ToolResult { Success = false, Error = "获取目标文档 subtype_format_mapping 失败" };
                }

                FtLog($"FetchSubtypeFormatMapping ok count={targetMapping.Count}");

                Dictionary<string, object> pageLayoutPayload;
                using (new FtStep("FetchSectionPageLayout"))
                {
                    pageLayoutPayload = await FetchSectionPageLayoutAsync(
                        docName, kbUuid, storageUuid).ConfigureAwait(false);
                }

                FtLog($"FetchSectionPageLayout ok has_payload={pageLayoutPayload != null}");

                List<Dictionary<string, object>> localMapping;
                List<Dictionary<string, object>> merged;
                Dictionary<string, object> remoteBodyFormat;
                using (new FtStep("MergeMapping"))
                {
                    localMapping = DocumentState.SubtypeFormatMapping;
                    Dictionary<string, Dictionary<string, object>> remoteFormatByDetailedSubtype =
                        BuildRemoteFormatByDetailedSubtypeFirstOccurrence(targetMapping);

                    merged = MergeLocalFormatsFromRemoteByDetailedSubtype(localMapping, remoteFormatByDetailedSubtype);
                    remoteBodyFormat = BuildRemoteBodyFormat(targetMapping);
                    merged = MergeLocalBodyFromRemote(merged, remoteBodyFormat);
                    DocumentState.SetSubtypeFormatMapping(merged);
                }

                FtLog(
                    $"MergeMapping ok local={(localMapping?.Count ?? 0)} merged={merged.Count} " +
                    $"remote_body={(remoteBodyFormat != null && remoteBodyFormat.Count > 0)}");

                var warnings = new List<string>();
                string contentMode = (pageLayoutContentMode ?? "auto").Trim();
                if (string.IsNullOrEmpty(contentMode))
                {
                    contentMode = "auto";
                }

                bool pageLayoutApplied = false;
                bool pageSetupApplied = false;
                int sectionsProcessed = 0;
                int headersFootersUpdated = 0;
                int hfContentFullCount = 0;
                int hfContentFormatOnlyCount = 0;
                var hfContentModes = new List<string>();
                var pageSetupWarnings = new List<string>();

                if (TryGetPageLayoutMode(pageLayoutPayload, out Dictionary<string, object> layoutMode))
                {
                    using (new FtStep(
                        "ApplyPageLayout",
                        $"content_mode={contentMode} apply_page_setup={applyPageSetup}"))
                    {
                        PageLayoutApplyResult plResult = PageLayoutApplyHelper.ApplyModeToDocument(
                            doc,
                            layoutMode,
                            contentMode,
                            new PageLayoutApplyOptions
                            {
                                ApplyPageSetup = applyPageSetup,
                                ApplyHeadersFooters = true,
                            });
                        pageLayoutApplied = plResult.Success;
                        pageSetupApplied = plResult.PageSetupApplied;
                        sectionsProcessed = plResult.SectionsProcessed;
                        headersFootersUpdated = plResult.HeadersFootersUpdated;
                        hfContentFullCount = plResult.HfContentFullCount;
                        hfContentFormatOnlyCount = plResult.HfContentFormatOnlyCount;
                        if (plResult.HfContentModes != null && plResult.HfContentModes.Count > 0)
                        {
                            hfContentModes.AddRange(plResult.HfContentModes);
                        }

                        if (plResult.Warnings != null && plResult.Warnings.Count > 0)
                        {
                            warnings.AddRange(plResult.Warnings);
                            foreach (string w in plResult.Warnings)
                            {
                                if (w != null
                                    && (w.IndexOf("page_setup", StringComparison.OrdinalIgnoreCase) >= 0
                                        || w.IndexOf("PageSetup", StringComparison.OrdinalIgnoreCase) >= 0
                                        || w.IndexOf("printer", StringComparison.OrdinalIgnoreCase) >= 0))
                                {
                                    pageSetupWarnings.Add(w);
                                }
                            }
                        }

                        if (applyPageSetup && !plResult.PrinterReady)
                        {
                            // D6：粗迁无打印机时跳过 page_setup，整体仍可成功
                            pageSetupApplied = false;
                        }
                    }

                    FtLog(
                        $"ApplyPageLayout result success={pageLayoutApplied} page_setup_applied={pageSetupApplied} " +
                        $"sections={sectionsProcessed} hf_updated={headersFootersUpdated}");
                }
                else
                {
                    warnings.Add("KB 无 section_page_layout.mode，已跳过页布局");
                    FtLog("ApplyPageLayout skipped (no mode)");
                }

                // 1. 整篇先套目标 KB 的 body 格式（一次 Content，替代逐句 Find+Apply）
                bool bodyApplied = false;
                if (remoteBodyFormat != null && remoteBodyFormat.Count > 0)
                {
                    using (new FtStep("ApplyBodyFormatToWholeDocument", $"keys={remoteBodyFormat.Count}"))
                    {
                        bodyApplied = ApplyBodyFormatToWholeDocument(doc, remoteBodyFormat);
                    }

                    if (!bodyApplied)
                    {
                        warnings.Add("整篇 body 格式应用失败");
                    }

                    FtLog($"ApplyBodyFormatToWholeDocument result={bodyApplied}");
                }
                else
                {
                    warnings.Add("mapping 无 body 条目，正文未应用");
                    FtLog("ApplyBodyFormatToWholeDocument skipped (no body)");
                }

                // 2. 再按 detailed_subtype 对齐标题（覆盖整篇 body 中的标题行）
                int headingsFormatted;
                using (new FtStep("ApplyFormatsToHeadingsByDetailedSubtype"))
                {
                    headingsFormatted = ApplyFormatsToHeadingsByDetailedSubtype(doc, merged);
                }

                FtLog($"ApplyFormatsToHeadingsByDetailedSubtype count={headingsFormatted}");

                // 3. 表格 standard 三字体：默认关闭（body 整篇 apply 已含表内文字）
                int tablesFormatted = 0;
                bool tableFormatStandardApplied = false;
#pragma warning disable CS0162 // ApplyTableFormatStandardInCoarseTransfer 为 false 时保留分支供后续启用
                if (ApplyTableFormatStandardInCoarseTransfer)
                {
                    Dictionary<string, object> tableStandard;
                    using (new FtStep("FetchTableFormatStandard"))
                    {
                        tableStandard = await FetchTableFormatStandardAsync(
                            docName, kbUuid, storageUuid).ConfigureAwait(false);
                    }

                    using (new FtStep("ApplyKbTableFormatStandard"))
                    {
                        TableStandardApplyResult tableResult = ApplyKbTableFormatStandardToDocument(doc, tableStandard);
                        tablesFormatted = tableResult.TablesFormatted;
                        tableFormatStandardApplied = tableResult.Applied;
                        if (tableResult.Warnings != null && tableResult.Warnings.Count > 0)
                        {
                            warnings.AddRange(tableResult.Warnings);
                        }
                    }
                }
#pragma warning restore CS0162
                else
                {
                    warnings.Add("table_format_standard_skipped: 粗迁暂不单独套用表字体，表内文字随 body 整篇 apply");
                    FtLog("ApplyKbTableFormatStandard skipped (flag=false)");
                }

                var data = new Dictionary<string, object>
                {
                    { "target_document_name", docName },
                    { "target_knowledge_base_uuid", kbUuid },
                    { "target_storage_doc_uuid", storageUuid },
                    { "headings_formatted", headingsFormatted },
                    { "body_applied", bodyApplied },
                    { "body_apply_mode", bodyApplied ? "whole_document" : "none" },
                    { "body_sentences_formatted", bodyApplied ? 1 : 0 },
                    { "page_layout_applied", pageLayoutApplied },
                    { "page_setup_applied", pageSetupApplied },
                    { "apply_page_setup", applyPageSetup },
                    { "sections_processed", sectionsProcessed },
                    { "headers_footers_updated", headersFootersUpdated },
                    { "hf_content_full_count", hfContentFullCount },
                    { "hf_content_format_only_count", hfContentFormatOnlyCount },
                    { "hf_content_modes", hfContentModes },
                    { "page_layout_content_mode", contentMode },
                    { "tables_formatted", tablesFormatted },
                    { "table_format_standard_applied", tableFormatStandardApplied }
                };
                if (warnings.Count > 0)
                {
                    data["warnings"] = warnings;
                }

                if (pageSetupWarnings.Count > 0)
                {
                    data["page_setup_warnings"] = pageSetupWarnings;
                }

                if (totalSw != null)
                {
                    totalSw.Stop();
                    FtLog($"END success total_elapsed_ms={totalSw.ElapsedMilliseconds} headings={headingsFormatted} body={bodyApplied}");
                }
                else
                {
                    FtLog($"END success headings={headingsFormatted} body={bodyApplied}");
                }

                return new ToolResult
                {
                    Success = true,
                    Data = data
                };
            }
            catch (Exception ex)
            {
                if (totalSw != null)
                {
                    totalSw.Stop();
                    FtLog($"END exception after_ms={totalSw.ElapsedMilliseconds}: {ex.GetType().Name}: {ex.Message}");
                }
                else
                {
                    FtLog($"END exception: {ex.GetType().Name}: {ex.Message}");
                }

                return new ToolResult { Success = false, Error = $"格式转移失败: {ex.Message}" };
            }
        }

        /// <summary>
        /// display→pool→format（R8）：复用 <see cref="WordDocumentExtractor.ProcessDocument"/>，不单独 readText 抽样。
        /// </summary>
        private static void RunReadAndFormatExtractionForTransfer(Word.Document document)
        {
            ProcessDocumentOptions options = ProcessDocumentOptions.ForFormatTransfer();
            // format_transfer 开启时打开 ProcessDocument 详细日志，便于对照卡在读入/格式提取哪一步
            if (FtDebugEnabled)
            {
                options.VerboseDebug = true;
                FtLog("ProcessDocument VerboseDebug=true (format_transfer)");
            }

            WordDocumentExtractor.ProcessDocument(document, options);
        }

        public static async Task<List<Dictionary<string, object>>> FetchSubtypeFormatMappingAsync(
            string documentName,
            string knowledgeBaseUuid,
            string storageDocUuid)
        {
            string baseUrl = (ConfigManager.Config.Api.BaseUrl ?? "").TrimEnd('/');
            if (string.IsNullOrEmpty(baseUrl))
            {
                FtLog("FetchSubtypeFormatMapping Api.BaseUrl 为空");
                System.Diagnostics.Debug.WriteLine("[FormatTransferHelper] Api.BaseUrl 为空");
                return null;
            }

            string url;
            if (!string.IsNullOrEmpty(storageDocUuid))
            {
                url =
                    $"{baseUrl}/knowledge/document/subtype_format_mapping_by_filename?storage_doc_uuid={Uri.EscapeDataString(storageDocUuid)}";
            }
            else
            {
                url =
                    $"{baseUrl}/knowledge/document/subtype_format_mapping_by_filename?document_name={Uri.EscapeDataString(documentName)}&knowledge_base_uuid={Uri.EscapeDataString(knowledgeBaseUuid)}";
            }

            FtLog($"HTTP GET subtype_format_mapping url={url}");
            var httpSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
            BackendApiClient.JsonResult result = await BackendApiClient.GetAuthenticatedAsync(url, retryOnUnauthorized: true, timeout: TimeSpan.FromSeconds(120))
                .ConfigureAwait(false);
            if (httpSw != null)
            {
                httpSw.Stop();
                FtLog($"HTTP GET subtype_format_mapping done status={(int)result.StatusCode} success={result.Success} elapsed_ms={httpSw.ElapsedMilliseconds}");
            }

            if (!result.Success)
            {
                FtLog($"FetchSubtypeFormatMapping fail status={(int)result.StatusCode}");
                System.Diagnostics.Debug.WriteLine($"[FormatTransferHelper] GET 失败 {(int)result.StatusCode}: {result.Body}");
                return null;
            }

            try
            {
                var root = JObject.Parse(result.Body);
                if (root["success"]?.Value<bool>() != true)
                {
                    FtLog("FetchSubtypeFormatMapping response success!=true");
                    return null;
                }

                JToken data = root["data"]?["subtype_format_mapping"];
                if (data == null || data.Type == JTokenType.Null)
                {
                    return new List<Dictionary<string, object>>();
                }

                return JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(data.ToString());
            }
            catch (Exception ex)
            {
                FtLog($"FetchSubtypeFormatMapping parse error: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[FormatTransferHelper] 解析响应失败: {ex.Message}");
                return null;
            }
        }

        public static async Task<Dictionary<string, object>> FetchParaFormatClustersAsync(
            string documentName,
            string knowledgeBaseUuid,
            string storageDocUuid)
        {
            string baseUrl = (ConfigManager.Config.Api.BaseUrl ?? "").TrimEnd('/');
            if (string.IsNullOrEmpty(baseUrl))
            {
                System.Diagnostics.Debug.WriteLine("[FormatTransferHelper] Api.BaseUrl 为空");
                return null;
            }

            string url;
            if (!string.IsNullOrEmpty(storageDocUuid))
            {
                url =
                    $"{baseUrl}/knowledge/document/para_format_clusters_by_filename?storage_doc_uuid={Uri.EscapeDataString(storageDocUuid)}";
            }
            else
            {
                url =
                    $"{baseUrl}/knowledge/document/para_format_clusters_by_filename?document_name={Uri.EscapeDataString(documentName)}&knowledge_base_uuid={Uri.EscapeDataString(knowledgeBaseUuid)}";
            }

            BackendApiClient.JsonResult result = await BackendApiClient.GetAuthenticatedAsync(
                    url, retryOnUnauthorized: true, timeout: TimeSpan.FromSeconds(120))
                .ConfigureAwait(false);

            if (!result.Success)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[FormatTransferHelper] GET para_format_clusters 失败 {(int)result.StatusCode}: {result.Body}");
                return null;
            }

            try
            {
                var root = JObject.Parse(result.Body);
                if (root["success"]?.Value<bool>() != true)
                {
                    return null;
                }

                JToken data = root["data"]?["para_format_clusters"];
                if (data == null || data.Type == JTokenType.Null)
                {
                    return new Dictionary<string, object>();
                }

                return JsonConvert.DeserializeObject<Dictionary<string, object>>(data.ToString())
                    ?? new Dictionary<string, object>();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[FormatTransferHelper] 解析 para_format_clusters 失败: {ex.Message}");
                return null;
            }
        }

        public static async Task<Dictionary<string, object>> FetchSectionPageLayoutAsync(
            string documentName,
            string knowledgeBaseUuid,
            string storageDocUuid)
        {
            string baseUrl = (ConfigManager.Config.Api.BaseUrl ?? "").TrimEnd('/');
            if (string.IsNullOrEmpty(baseUrl))
            {
                System.Diagnostics.Debug.WriteLine("[FormatTransferHelper] Api.BaseUrl 为空");
                return null;
            }

            string url;
            if (!string.IsNullOrEmpty(storageDocUuid))
            {
                url =
                    $"{baseUrl}/knowledge/document/section_page_layout_by_filename?storage_doc_uuid={Uri.EscapeDataString(storageDocUuid)}";
            }
            else
            {
                url =
                    $"{baseUrl}/knowledge/document/section_page_layout_by_filename?document_name={Uri.EscapeDataString(documentName)}&knowledge_base_uuid={Uri.EscapeDataString(knowledgeBaseUuid)}";
            }

            FtLog($"HTTP GET section_page_layout url={url}");
            var httpSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
            BackendApiClient.JsonResult result = await BackendApiClient.GetAuthenticatedAsync(
                    url, retryOnUnauthorized: true, timeout: TimeSpan.FromSeconds(120))
                .ConfigureAwait(false);
            if (httpSw != null)
            {
                httpSw.Stop();
                FtLog($"HTTP GET section_page_layout done status={(int)result.StatusCode} success={result.Success} elapsed_ms={httpSw.ElapsedMilliseconds}");
            }

            if (!result.Success)
            {
                if ((int)result.StatusCode == 404)
                {
                    FtLog("section_page_layout 404，跳过页布局");
                    System.Diagnostics.Debug.WriteLine("[FormatTransferHelper] section_page_layout 404，跳过页布局");
                    return null;
                }

                FtLog($"FetchSectionPageLayout fail status={(int)result.StatusCode}");
                System.Diagnostics.Debug.WriteLine(
                    $"[FormatTransferHelper] GET section_page_layout 失败 {(int)result.StatusCode}: {result.Body}");
                return null;
            }

            try
            {
                var root = JObject.Parse(result.Body);
                if (root["success"]?.Value<bool>() != true)
                {
                    FtLog("FetchSectionPageLayout response success!=true");
                    return null;
                }

                JToken data = root["data"]?["section_page_layout"];
                if (data == null || data.Type == JTokenType.Null)
                {
                    return null;
                }

                return JsonConvert.DeserializeObject<Dictionary<string, object>>(data.ToString());
            }
            catch (Exception ex)
            {
                FtLog($"FetchSectionPageLayout parse error: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[FormatTransferHelper] 解析 section_page_layout 失败: {ex.Message}");
                return null;
            }
        }

        public sealed class TableStandardApplyResult
        {
            public bool Applied { get; set; }
            public int TablesFormatted { get; set; }
            public List<string> Warnings { get; set; } = new List<string>();
        }

        public static async Task<Dictionary<string, object>> FetchTableFormatStandardAsync(
            string documentName,
            string knowledgeBaseUuid,
            string storageDocUuid)
        {
            string baseUrl = (ConfigManager.Config.Api.BaseUrl ?? "").TrimEnd('/');
            if (string.IsNullOrEmpty(baseUrl))
            {
                System.Diagnostics.Debug.WriteLine("[FormatTransferHelper] Api.BaseUrl 为空");
                return null;
            }

            string url;
            if (!string.IsNullOrEmpty(storageDocUuid))
            {
                url =
                    $"{baseUrl}/knowledge/document/table_format_standard_by_filename?storage_doc_uuid={Uri.EscapeDataString(storageDocUuid)}";
            }
            else
            {
                url =
                    $"{baseUrl}/knowledge/document/table_format_standard_by_filename?document_name={Uri.EscapeDataString(documentName)}&knowledge_base_uuid={Uri.EscapeDataString(knowledgeBaseUuid)}";
            }

            BackendApiClient.JsonResult result = await BackendApiClient.GetAuthenticatedAsync(
                    url, retryOnUnauthorized: true, timeout: TimeSpan.FromSeconds(120))
                .ConfigureAwait(false);

            if (!result.Success)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[FormatTransferHelper] GET table_format_standard 失败 {(int)result.StatusCode}: {result.Body}");
                return null;
            }

            try
            {
                var root = JObject.Parse(result.Body);
                if (root["success"]?.Value<bool>() != true)
                {
                    return null;
                }

                JToken data = root["data"];
                if (data == null || data.Type == JTokenType.Null)
                {
                    return null;
                }

                bool missing = data["table_format_standard_missing"]?.Value<bool>() ?? false;
                if (missing)
                {
                    var stub = new Dictionary<string, object>
                    {
                        { "_missing", true },
                        { "_warning", data["table_format_standard_warning"]?.ToString() ?? "table_format_standard 不存在" },
                    };
                    return stub;
                }

                JToken standard = data["table_format_standard"];
                if (standard == null || standard.Type == JTokenType.Null)
                {
                    return null;
                }

                return JsonConvert.DeserializeObject<Dictionary<string, object>>(standard.ToString());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[FormatTransferHelper] 解析 table_format_standard 失败: {ex.Message}");
                return null;
            }
        }

        public static string BuildMinimalTableStandardXml(Dictionary<string, object> standard)
        {
            if (standard == null || standard.Count == 0)
            {
                return null;
            }

            if (standard.TryGetValue("_missing", out object missingFlag)
                && missingFlag is bool b
                && b)
            {
                return null;
            }

            string fontName = GetDictString(standard, "table_font_name");
            string fontColor = GetDictString(standard, "table_font_color");
            TryGetDouble(standard, "table_font_size", out double fontSize);

            if (string.IsNullOrWhiteSpace(fontName)
                && string.IsNullOrWhiteSpace(fontColor)
                && fontSize <= 0)
            {
                return null;
            }

            var root = new XElement("TableConfig");
            var general = new XElement("General");
            if (!string.IsNullOrWhiteSpace(fontColor))
            {
                general.Add(new XElement("TableFontColor", fontColor.Trim()));
            }
            if (!string.IsNullOrWhiteSpace(fontName))
            {
                general.Add(new XElement("TableFontName", fontName.Trim()));
            }
            if (fontSize > 0)
            {
                general.Add(
                    new XElement(
                        "TableFontSize",
                        fontSize.ToString(CultureInfo.InvariantCulture)));
            }
            root.Add(general);
            return root.ToString();
        }

        public static TableStandardApplyResult ApplyKbTableFormatStandardToDocument(
            Word.Document doc,
            Dictionary<string, object> standard)
        {
            var result = new TableStandardApplyResult();
            if (doc == null)
            {
                result.Warnings.Add("table_format_standard_skipped: 文档不可用");
                return result;
            }

            if (standard == null)
            {
                result.Warnings.Add(
                    "table_format_standard_skipped: 拉取 table_format_standard 失败，表字体未改");
                return result;
            }

            if (standard.TryGetValue("_missing", out object missingFlag)
                && missingFlag is bool b
                && b)
            {
                string msg = standard.TryGetValue("_warning", out object w) && w != null
                    ? w.ToString()
                    : "table_format_standard.xml 不存在";
                result.Warnings.Add($"table_format_standard_skipped: {msg}");
                return result;
            }

            string xml = BuildMinimalTableStandardXml(standard);
            if (string.IsNullOrWhiteSpace(xml))
            {
                result.Warnings.Add("table_format_standard_skipped: standard 无有效三字体字段");
                return result;
            }

            DocumentState.BindAndActivate(doc);
            List<string> tableIds = DocumentState.GetTableIdOrder();
            if (tableIds == null || tableIds.Count == 0)
            {
                result.Warnings.Add("table_format_standard_skipped: 当前稿无表格（T_ 为空）");
                return result;
            }

            int formatted = 0;
            foreach (string tableId in tableIds)
            {
                if (string.IsNullOrWhiteSpace(tableId))
                {
                    continue;
                }

                TableConfigApplyResult applyResult = TableConfigApplyHelper.ApplyTableConfigToExistingTable(
                    doc, tableId.Trim(), xml);
                if (!applyResult.Success)
                {
                    result.Warnings.Add(
                        $"table_format_apply_failed:{tableId}:{applyResult.Error ?? "unknown"}");
                    continue;
                }

                formatted++;
                if (applyResult.Warnings != null && applyResult.Warnings.Count > 0)
                {
                    foreach (string w in applyResult.Warnings)
                    {
                        result.Warnings.Add($"{tableId}:{w}");
                    }
                }
            }

            result.TablesFormatted = formatted;
            result.Applied = formatted > 0;
            System.Diagnostics.Debug.WriteLine(
                $"[FormatTransferHelper] 表格 standard 已 apply {formatted}/{tableIds.Count} 张表");
            return result;
        }

        private static bool TryGetDouble(Dictionary<string, object> dict, string key, out double value)
        {
            value = 0;
            if (dict == null || !dict.TryGetValue(key, out object o) || o == null)
            {
                return false;
            }

            if (o is double d)
            {
                value = d;
                return true;
            }

            if (o is float f)
            {
                value = f;
                return true;
            }

            if (o is int i)
            {
                value = i;
                return true;
            }

            if (o is long l)
            {
                value = l;
                return true;
            }

            return double.TryParse(
                o.ToString(),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out value);
        }

        public static bool TryGetPageLayoutMode(
            Dictionary<string, object> pageLayoutPayload,
            out Dictionary<string, object> mode)
        {
            mode = null;
            if (pageLayoutPayload == null || pageLayoutPayload.Count == 0)
            {
                return false;
            }

            if (!pageLayoutPayload.TryGetValue("mode", out object modeObj) || modeObj == null)
            {
                return false;
            }

            mode = NormalizeFormatDict(modeObj);
            return mode != null && mode.Count > 0;
        }

        /// <summary>
        /// 远程 mapping 中按 detailed_subtype 索引：每个 subtype 只取第一次出现的 heading format（P9 无 level）。
        /// </summary>
        private static Dictionary<string, Dictionary<string, object>> BuildRemoteFormatByDetailedSubtypeFirstOccurrence(
            List<Dictionary<string, object>> targetMapping)
        {
            var map = new Dictionary<string, Dictionary<string, object>>(StringComparer.Ordinal);
            foreach (Dictionary<string, object> item in targetMapping)
            {
                if (item == null || GetDictString(item, "type") != "heading")
                {
                    continue;
                }

                string dst = GetDictString(item, "detailed_subtype");
                if (string.IsNullOrEmpty(dst) || map.ContainsKey(dst))
                {
                    continue;
                }

                Dictionary<string, object> fmt = NormalizeFormatDict(GetDictObject(item, "format"));
                if (fmt != null && fmt.Count > 0)
                {
                    map[dst] = CloneFormatDict(fmt);
                }
            }

            return map;
        }

        /// <summary>
        /// 本地 heading 若远程存在同 detailed_subtype 的 format，则替换 format。
        /// </summary>
        private static List<Dictionary<string, object>> MergeLocalFormatsFromRemoteByDetailedSubtype(
            List<Dictionary<string, object>> localMapping,
            Dictionary<string, Dictionary<string, object>> remoteFormatByDetailedSubtype)
        {
            var merged = new List<Dictionary<string, object>>();
            foreach (Dictionary<string, object> item in localMapping)
            {
                Dictionary<string, object> copy = CloneMappingItem(item);
                if (GetDictString(copy, "type") == "heading")
                {
                    string dst = GetDictString(copy, "detailed_subtype");
                    if (!string.IsNullOrEmpty(dst)
                        && remoteFormatByDetailedSubtype.TryGetValue(dst, out Dictionary<string, object> remoteFmt))
                    {
                        copy["format"] = CloneFormatDict(remoteFmt);
                    }
                }

                merged.Add(copy);
            }

            return merged;
        }

        private static int ApplyFormatsToHeadingsByDetailedSubtype(
            Word.Document doc,
            List<Dictionary<string, object>> mergedMapping)
        {
            var dstToFormat = new Dictionary<string, Dictionary<string, object>>(StringComparer.Ordinal);
            foreach (Dictionary<string, object> item in mergedMapping)
            {
                if (item == null || GetDictString(item, "type") != "heading")
                {
                    continue;
                }

                string dst = GetDictString(item, "detailed_subtype");
                if (string.IsNullOrEmpty(dst))
                {
                    continue;
                }

                Dictionary<string, object> fmt = NormalizeFormatDict(GetDictObject(item, "format"));
                if (fmt != null && fmt.Count > 0)
                {
                    dstToFormat[dst] = fmt;
                }
            }

            int count = 0;
            List<Dictionary<string, object>> headings = DocumentState.RecognizedHeadingsInDocumentOrder
                ?? new List<Dictionary<string, object>>();
            int headingTotal = headings.Count;
            FtLog($"headings apply start recognized={headingTotal} format_keys={dstToFormat.Count}");

            int idx = 0;
            foreach (Dictionary<string, object> h in headings)
            {
                idx++;
                if (h == null)
                {
                    continue;
                }

                string dst = GetDictString(h, "detailed_subtype");
                if (string.IsNullOrEmpty(dst) || !dstToFormat.TryGetValue(dst, out Dictionary<string, object> format))
                {
                    continue;
                }

                string text = GetDictString(h, "text");
                if (string.IsNullOrEmpty(text))
                {
                    continue;
                }

                string preview = text.Length > 40 ? text.Substring(0, 40) + "…" : text;
                preview = preview.Replace("\r", "\\r").Replace("\n", "\\n");
                FtLog($"heading[{idx}/{headingTotal}] FIND dst={dst} text={preview}");

                var findSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                List<Word.Range> ranges = WordRangeFinder.FindAllInDocument(doc, text, "FormatTransfer");
                if (findSw != null)
                {
                    findSw.Stop();
                    FtLog($"heading[{idx}/{headingTotal}] FIND done hits={ranges?.Count ?? 0} elapsed_ms={findSw.ElapsedMilliseconds}");
                }

                if (ranges == null || ranges.Count == 0)
                {
                    FtLog($"heading[{idx}/{headingTotal}] miss dst={dst}");
                    System.Diagnostics.Debug.WriteLine($"[FormatTransferHelper] 未找到标题文本: {dst}");
                    continue;
                }

                Dictionary<string, object> snapshot = FormatInheritHelper.BuildSnapshotFromCharFormat(format);
                if (snapshot == null || snapshot.Count == 0)
                {
                    continue;
                }

                int hitIndex = 0;
                foreach (Word.Range range in ranges)
                {
                    if (range == null)
                    {
                        continue;
                    }

                    hitIndex++;
                    FtLog($"heading[{idx}/{headingTotal}] APPLY dst={dst} hit={hitIndex}/{ranges.Count}");
                    var applySw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                    FormatInheritHelper.ApplySnapshot(range, snapshot, charFormatOnly: true);
                    if (applySw != null)
                    {
                        applySw.Stop();
                        FtLog($"heading[{idx}/{headingTotal}] APPLY hit={hitIndex} done elapsed_ms={applySw.ElapsedMilliseconds}");
                    }

                    count++;
                }
            }

            FtLog($"headings apply end formatted={count}");
            return count;
        }

        /// <summary>
        /// 对 document.Content 一次性应用目标 KB body 字符格式（charFormatOnly，不改段落样式名；不写底纹/背景色）。
        /// </summary>
        private static bool ApplyBodyFormatToWholeDocument(Word.Document doc, Dictionary<string, object> bodyFormat)
        {
            if (doc == null || bodyFormat == null || bodyFormat.Count == 0)
            {
                return false;
            }

            try
            {
                FtLog("body BuildSnapshotFromCharFormat…");
                Dictionary<string, object> snapshot = FormatInheritHelper.BuildSnapshotFromCharFormat(
                    bodyFormat,
                    includeBackgroundColor: false);
                if (snapshot == null || snapshot.Count == 0)
                {
                    FtLog("body snapshot empty");
                    return false;
                }

                FtLog($"body ApplySnapshot on doc.Content keys={snapshot.Count}…");
                var applySw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                FormatInheritHelper.ApplySnapshot(doc.Content, snapshot, charFormatOnly: true);
                if (applySw != null)
                {
                    applySw.Stop();
                    FtLog($"body ApplySnapshot done elapsed_ms={applySw.ElapsedMilliseconds}");
                }

                System.Diagnostics.Debug.WriteLine("[FormatTransferHelper] 已整篇应用 body 格式 (whole_document)");
                return true;
            }
            catch (Exception ex)
            {
                FtLog($"body ApplySnapshot exception: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[FormatTransferHelper] ApplyBodyFormatToWholeDocument: {ex.Message}");
                return false;
            }
        }

        private static Dictionary<string, object> BuildRemoteBodyFormat(List<Dictionary<string, object>> targetMapping)
        {
            if (targetMapping == null)
            {
                return null;
            }

            foreach (Dictionary<string, object> item in targetMapping)
            {
                if (item == null)
                {
                    continue;
                }

                string dst = GetDictString(item, "detailed_subtype");
                if (!string.Equals(dst, "body", StringComparison.OrdinalIgnoreCase)
                    && GetDictString(item, "type") != "body")
                {
                    continue;
                }

                Dictionary<string, object> fmt = NormalizeFormatDict(GetDictObject(item, "format"));
                if (fmt != null && fmt.Count > 0)
                {
                    return fmt;
                }
            }

            return null;
        }

        private static List<Dictionary<string, object>> MergeLocalBodyFromRemote(
            List<Dictionary<string, object>> localMapping,
            Dictionary<string, object> remoteBodyFormat)
        {
            if (remoteBodyFormat == null || remoteBodyFormat.Count == 0)
            {
                return localMapping;
            }

            var merged = new List<Dictionary<string, object>>();
            bool foundBody = false;
            foreach (Dictionary<string, object> item in localMapping)
            {
                Dictionary<string, object> copy = CloneMappingItem(item);
                if (GetDictString(copy, "type") == "body"
                    || string.Equals(GetDictString(copy, "detailed_subtype"), "body", StringComparison.OrdinalIgnoreCase))
                {
                    copy["type"] = "body";
                    copy["detailed_subtype"] = "body";
                    copy["format"] = CloneFormatDict(remoteBodyFormat);
                    foundBody = true;
                }

                merged.Add(copy);
            }

            if (!foundBody)
            {
                merged.Add(new Dictionary<string, object>
                {
                    { "type", "body" },
                    { "subtype", "" },
                    { "detailed_subtype", "body" },
                    { "language", "" },
                    { "format", CloneFormatDict(remoteBodyFormat) }
                });
            }

            return merged;
        }

        private static Dictionary<string, object> NormalizeFormatDict(object fmtObj)
        {
            if (fmtObj == null)
            {
                return new Dictionary<string, object>();
            }

            if (fmtObj is Dictionary<string, object> d)
            {
                return d;
            }

            if (fmtObj is JObject jo)
            {
                return jo.ToObject<Dictionary<string, object>>() ?? new Dictionary<string, object>();
            }

            try
            {
                return JsonConvert.DeserializeObject<Dictionary<string, object>>(fmtObj.ToString());
            }
            catch
            {
                return new Dictionary<string, object>();
            }
        }

        private static Dictionary<string, object> CloneFormatDict(Dictionary<string, object> src)
        {
            var d = new Dictionary<string, object>();
            if (src == null)
            {
                return d;
            }

            foreach (var kvp in src)
            {
                d[kvp.Key] = kvp.Value;
            }

            return d;
        }

        private static Dictionary<string, object> CloneMappingItem(Dictionary<string, object> item)
        {
            var copy = new Dictionary<string, object>();
            if (item == null)
            {
                return copy;
            }

            foreach (var kvp in item)
            {
                if (kvp.Value is Dictionary<string, object> nested)
                {
                    copy[kvp.Key] = CloneFormatDict(nested);
                }
                else
                {
                    copy[kvp.Key] = kvp.Value;
                }
            }

            return copy;
        }

        private static bool TryGetInt(Dictionary<string, object> dict, string key, out int value)
        {
            value = 0;
            if (dict == null || !dict.TryGetValue(key, out object o) || o == null)
            {
                return false;
            }

            if (o is int i)
            {
                value = i;
                return true;
            }

            if (o is long l)
            {
                value = (int)l;
                return true;
            }

            return int.TryParse(o.ToString(), out value);
        }

        private static string GetDictString(Dictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.TryGetValue(key, out object o) || o == null)
            {
                return "";
            }

            return o.ToString() ?? "";
        }

        private static object GetDictObject(Dictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.TryGetValue(key, out object o))
            {
                return null;
            }

            return o;
        }
    }
}
