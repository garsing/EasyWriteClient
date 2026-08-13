using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 应用页面设置：page_setup（边距/纸张）与/或页眉页脚（KB mode）。
    /// 文档触点经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_ApplyPageSetupTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_page_setup"] = async (args) =>
            {
                try
                {
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_apply_page_setup] resolve failed: {resolveError?.Error}; " +
                            $"arg.channel_id={ChannelContext.TryGetChannelIdFromParameters(args) ?? "(null)"}; " +
                            $"default={ChannelRegistry.DefaultChannelId ?? "(null)"}");
                        return Fail(resolveError?.Error ?? "无法解析文档渠道");
                    }

                    Word.Document doc = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_apply_page_setup] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    bool applyPageSetup = GetBoolArg(args, "apply_page_setup", true);
                    bool applyHeadersFooters = GetBoolArg(args, "apply_headers_footers", true);
                    if (!applyPageSetup && !applyHeadersFooters)
                    {
                        return Fail("apply_page_setup 与 apply_headers_footers 不能同时为 false");
                    }

                    string contentMode = GetStringArg(args, "page_layout_content_mode");
                    if (string.IsNullOrEmpty(contentMode))
                    {
                        contentMode = "auto";
                    }

                    int? sectionIndex = null;
                    if (args != null && args.TryGetValue("section_index", out object secObj) && secObj != null)
                    {
                        if (TryToInt(secObj, out int sec) && sec >= 1)
                        {
                            sectionIndex = sec;
                        }
                    }

                    string storageUuid = GetStringArg(args, "target_storage_doc_uuid");
                    string docName = GetStringArg(args, "target_document_name");
                    string kbUuid = GetStringArg(args, "target_knowledge_base_uuid");

                    Dictionary<string, object> mode = null;
                    if (!string.IsNullOrEmpty(storageUuid) || !string.IsNullOrEmpty(docName))
                    {
                        if (string.IsNullOrEmpty(storageUuid) && string.IsNullOrEmpty(kbUuid))
                        {
                            return Fail("按 document_name 定位时必须提供 target_knowledge_base_uuid");
                        }

                        var userService = UserService.Instance;
                        if (!userService.CheckLoginStatus())
                        {
                            return Fail("用户未登录，无法调用知识库接口");
                        }

                        Dictionary<string, object> payload = await FormatTransferHelper.FetchSectionPageLayoutAsync(
                                docName, kbUuid, storageUuid)
                            .ConfigureAwait(true);
                        if (!FormatTransferHelper.TryGetPageLayoutMode(payload, out mode) || mode == null)
                        {
                            return Fail("获取 KB section_page_layout.mode 失败或 mode 为空");
                        }
                    }
                    else
                    {
                        mode = new Dictionary<string, object>(StringComparer.Ordinal);
                    }

                    Dictionary<string, object> pageSetup = GetOrCreateDict(mode, "page_setup");
                    MergeExplicitPageSetupArgs(args, pageSetup);

                    if (applyHeadersFooters
                        && !mode.ContainsKey("headers")
                        && !mode.ContainsKey("footers")
                        && string.IsNullOrEmpty(storageUuid)
                        && string.IsNullOrEmpty(docName))
                    {
                        if (!applyPageSetup || pageSetup.Count == 0)
                        {
                            return Fail("未提供 KB 定位且无可用 page_setup / headers/footers");
                        }

                        // 仅显式边距：关闭 HF
                        applyHeadersFooters = false;
                    }

                    if (applyPageSetup && pageSetup.Count == 0 && applyHeadersFooters)
                    {
                        // 仅 HF：允许 page_setup 为空（开关可能仍从空 dict 读 false）
                    }
                    else if (applyPageSetup && pageSetup.Count == 0 && !applyHeadersFooters)
                    {
                        return Fail("未指定 page_setup 字段且未提供 KB mode");
                    }

                    mode["page_setup"] = pageSetup;

                    PageLayoutApplyResult plResult = PageLayoutApplyHelper.ApplyModeToDocument(
                        doc,
                        mode,
                        contentMode,
                        new PageLayoutApplyOptions
                        {
                            ApplyPageSetup = applyPageSetup,
                            ApplyHeadersFooters = applyHeadersFooters,
                            SectionIndex = sectionIndex,
                        });

                    bool success = EvaluateSuccess(
                        applyPageSetup,
                        applyHeadersFooters,
                        plResult);

                    var data = new Dictionary<string, object>
                    {
                        { "page_setup_applied", plResult.PageSetupApplied },
                        { "headers_footers_updated", plResult.HeadersFootersUpdated },
                        { "sections_processed", plResult.SectionsProcessed },
                        { "printer_switched", plResult.PrinterSwitched },
                        { "printer_ready", plResult.PrinterReady },
                        { "apply_page_setup", applyPageSetup },
                        { "apply_headers_footers", applyHeadersFooters },
                        { "page_layout_content_mode", contentMode },
                        { "hf_content_modes", plResult.HfContentModes ?? new List<string>() },
                        { "hf_content_full_count", plResult.HfContentFullCount },
                        { "hf_content_format_only_count", plResult.HfContentFormatOnlyCount },
                    };
                    if (sectionIndex.HasValue)
                    {
                        data["section_index"] = sectionIndex.Value;
                    }

                    if (plResult.Warnings != null && plResult.Warnings.Count > 0)
                    {
                        data["warnings"] = plResult.Warnings;
                    }

                    if (!success)
                    {
                        string err = "应用页面设置未成功";
                        if (plResult.Warnings != null && plResult.Warnings.Count > 0)
                        {
                            err = string.Join("; ", plResult.Warnings);
                        }

                        return new ToolResult { Success = false, Error = err, Data = data };
                    }

                    return new ToolResult { Success = true, Data = data };
                }
                catch (Exception ex)
                {
                    return Fail($"F_apply_page_setup 异常: {ex.Message}");
                }
            };
        }

        /// <summary>I6：无打印机但 HF 成功 → success=true；仅 PS 失败 → false。</summary>
        private static bool EvaluateSuccess(
            bool wantPs,
            bool wantHf,
            PageLayoutApplyResult pl)
        {
            if (pl == null)
            {
                return false;
            }

            bool hfOk = wantHf && (pl.HeadersFootersUpdated > 0 || pl.SectionsProcessed > 0);
            bool noPrinter = wantPs && !pl.PrinterReady;

            if (noPrinter)
            {
                return wantHf && hfOk;
            }

            if (wantPs && wantHf)
            {
                return pl.PageSetupApplied || hfOk;
            }

            if (wantPs)
            {
                return pl.PageSetupApplied;
            }

            return hfOk;
        }

        private static void MergeExplicitPageSetupArgs(
            Dictionary<string, object> args,
            Dictionary<string, object> pageSetup)
        {
            if (args == null || pageSetup == null)
            {
                return;
            }

            TryCopyNumber(args, pageSetup, "top_margin");
            TryCopyNumber(args, pageSetup, "bottom_margin");
            TryCopyNumber(args, pageSetup, "left_margin");
            TryCopyNumber(args, pageSetup, "right_margin");
            TryCopyNumber(args, pageSetup, "page_width");
            TryCopyNumber(args, pageSetup, "page_height");
            TryCopyNumber(args, pageSetup, "header_distance");
            TryCopyNumber(args, pageSetup, "footer_distance");

            string orientation = GetStringArg(args, "orientation");
            if (!string.IsNullOrEmpty(orientation))
            {
                pageSetup["orientation"] = orientation;
            }

            if (args.TryGetValue("different_first_page_header_footer", out object df) && df != null)
            {
                pageSetup["different_first_page_header_footer"] = GetBoolArg(args, "different_first_page_header_footer", false);
            }

            if (args.TryGetValue("odd_and_even_pages_header_footer", out object oe) && oe != null)
            {
                pageSetup["odd_and_even_pages_header_footer"] = GetBoolArg(args, "odd_and_even_pages_header_footer", false);
            }
        }

        private static void TryCopyNumber(
            Dictionary<string, object> args,
            Dictionary<string, object> pageSetup,
            string key)
        {
            if (!args.TryGetValue(key, out object o) || o == null)
            {
                return;
            }

            if (TryToDouble(o, out double d))
            {
                pageSetup[key] = d;
            }
        }

        private static Dictionary<string, object> GetOrCreateDict(
            Dictionary<string, object> mode,
            string key)
        {
            if (mode.TryGetValue(key, out object o) && o is Dictionary<string, object> d)
            {
                return d;
            }

            var created = new Dictionary<string, object>(StringComparer.Ordinal);
            mode[key] = created;
            return created;
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }

        private static string GetStringArg(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out object o) || o == null)
            {
                return "";
            }

            return o.ToString()?.Trim() ?? "";
        }

        private static bool GetBoolArg(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (args == null || !args.TryGetValue(key, out object o) || o == null)
            {
                return defaultValue;
            }

            if (o is bool b)
            {
                return b;
            }

            string s = o.ToString()?.Trim();
            if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase) || s == "1")
            {
                return true;
            }

            if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase) || s == "0")
            {
                return false;
            }

            return defaultValue;
        }

        private static bool TryToInt(object o, out int value)
        {
            value = 0;
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

            return int.TryParse(o.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }

        private static bool TryToDouble(object o, out double value)
        {
            value = 0;
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
    }
}
