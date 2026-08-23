using System;
using System.Collections.Generic;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public enum FormatSourceKind
    {
        None = 0,
        Local = 1,
        KnowledgeBase = 2
    }

    public sealed class FormatSourceResolution
    {
        public FormatSourceKind Kind { get; set; }

        public string Error { get; set; }

        public bool Success
        {
            get { return Kind != FormatSourceKind.None && string.IsNullOrEmpty(Error); }
        }

        public Word.Document Document { get; set; }

        public string DocUuid { get; set; }

        public string ChannelId { get; set; }

        public InteropDocumentHandle Handle { get; set; }

        public string StorageDocUuid { get; set; }

        public string DocumentName { get; set; }

        public string KnowledgeBaseUuid { get; set; }

        public bool ForceRefresh { get; set; }

        public bool DisallowBackendApi
        {
            get { return Handle != null && Handle.DisallowBackendApi; }
        }
    }

    /// <summary>
    /// I1：格式源定位。渠道与知识库互斥；废止 target_*；源=目标失败。
    /// </summary>
    public static class FormatSourceResolver
    {
        private static readonly string[] ForbiddenTargetKeys =
        {
            "target_storage_doc_uuid",
            "target_document_name",
            "target_knowledge_base_uuid"
        };

        public static string GetArgString(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out object raw) || raw == null)
            {
                return "";
            }

            return raw.ToString()?.Trim() ?? "";
        }

        public static bool GetBoolArg(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (args == null || !args.TryGetValue(key, out object raw) || raw == null)
            {
                return defaultValue;
            }

            if (raw is bool b)
            {
                return b;
            }

            string s = raw.ToString()?.Trim();
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

        public static bool TryRejectTargetAlias(Dictionary<string, object> args, out ToolResult error)
        {
            error = null;
            if (args == null)
            {
                return true;
            }

            foreach (string key in ForbiddenTargetKeys)
            {
                if (args.TryGetValue(key, out object raw) && raw != null
                    && !string.IsNullOrWhiteSpace(raw.ToString()))
                {
                    error = Fail("已废止 " + key + "，请改用 storage_doc_uuid / document_name / knowledge_base_uuid");
                    return false;
                }
            }

            return true;
        }

        /// <param name="requireSource">false 时允许无源（如页设置仅手填边距）。</param>
        public static bool TryResolve(
            Dictionary<string, object> args,
            object wordApplication,
            string targetDocUuid,
            out FormatSourceResolution source,
            out ToolResult error,
            bool requireSource = true,
            bool activateSource = false)
        {
            source = new FormatSourceResolution
            {
                ForceRefresh = GetBoolArg(args, "force_refresh", false)
            };
            error = null;

            if (!TryRejectTargetAlias(args, out error))
            {
                return false;
            }

            string sourceChannelId = GetArgString(args, "source_channel_id");
            string storageUuid = GetArgString(args, "storage_doc_uuid");
            string documentName = GetArgString(args, "document_name");
            string kbUuid = GetArgString(args, "knowledge_base_uuid");

            bool hasLocal = !string.IsNullOrEmpty(sourceChannelId);
            bool hasKb = !string.IsNullOrEmpty(storageUuid) || !string.IsNullOrEmpty(documentName);

            if (hasLocal && hasKb)
            {
                error = Fail("source_channel_id 与知识库定位不能同时传");
                return false;
            }

            if (!hasLocal && !hasKb)
            {
                if (!requireSource)
                {
                    source.Kind = FormatSourceKind.None;
                    return true;
                }

                error = Fail("缺少格式源：请传 source_channel_id，或 storage_doc_uuid（或 document_name + knowledge_base_uuid）");
                return false;
            }

            if (hasKb)
            {
                if (string.IsNullOrEmpty(storageUuid) && string.IsNullOrEmpty(kbUuid))
                {
                    error = Fail("按 document_name 定位时必须提供 knowledge_base_uuid");
                    return false;
                }

                source.Kind = FormatSourceKind.KnowledgeBase;
                source.StorageDocUuid = storageUuid;
                source.DocumentName = documentName;
                source.KnowledgeBaseUuid = kbUuid;
                return true;
            }

            var sourceArgs = new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["channel_id"] = sourceChannelId
            };

            if (!DocumentHostAdapter.TryResolveInteropDocument(
                    sourceArgs,
                    wordApplication,
                    out InteropDocumentHandle handle,
                    out ToolResult resolveError,
                    activateDocument: activateSource))
            {
                error = resolveError ?? Fail("无法解析 source_channel_id，请先打开模板文档");
                if (error != null && !string.IsNullOrEmpty(error.Error)
                    && error.Error.IndexOf("未知 channel_id", StringComparison.Ordinal) >= 0)
                {
                    error = Fail("无效 source_channel_id，请先打开模板文档");
                }

                return false;
            }

            string sourceUuid = handle.Context?.DocUuid ?? "";
            if (!string.IsNullOrEmpty(targetDocUuid)
                && !string.IsNullOrEmpty(sourceUuid)
                && string.Equals(sourceUuid, targetDocUuid, StringComparison.Ordinal))
            {
                error = Fail("格式源与目标是同一份文档，不能自迁");
                return false;
            }

            source.Kind = FormatSourceKind.Local;
            source.Handle = handle;
            source.Document = handle.Document;
            source.DocUuid = sourceUuid;
            source.ChannelId = handle.ChannelId;
            return true;
        }

        private static ToolResult Fail(string message)
        {
            return new ToolResult { Success = false, Error = message };
        }
    }
}
