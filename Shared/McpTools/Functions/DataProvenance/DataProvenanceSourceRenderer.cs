using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class DataProvenanceSourceRenderer
    {
        public static async Task<(RenderedProvenanceItem Item, string Error)> RenderItemAsync(
            ProvenanceItemInput input,
            int itemIndex)
        {
            if (input?.Anchor == null)
            {
                return (null, "缺少 anchor");
            }

            var rendered = new RenderedProvenanceItem
            {
                Anchor = input.Anchor,
                Kind = DataProvenanceParamValidator.ResolveKind(input.Anchor.Code),
            };

            try
            {
                if (rendered.Kind == ProvenanceAnchorKind.Sentence)
                {
                    string line = await RenderSourceAsync(input.Source).ConfigureAwait(false);
                    rendered.RenderedLineForReference = NormalizeRenderedLine(line);
                }
                else
                {
                    foreach (ProvenanceSourceInput src in input.Sources)
                    {
                        string line = await RenderSourceAsync(src).ConfigureAwait(false);
                        rendered.RenderedObjectLines.Add(NormalizeRenderedLine(line));
                    }
                }
            }
            catch (Exception ex)
            {
                return (null, ex.Message);
            }

            foreach (string line in EnumerateLines(rendered))
            {
                if (line != null && line.Length > DataProvenanceParamValidator.MaxRenderedLineLength)
                {
                    return (null, $"渲染后说明超过 {DataProvenanceParamValidator.MaxRenderedLineLength} 字");
                }
            }

            return (rendered, null);
        }

        public static ProvenanceValidationResult ValidateSentenceAnchor(
            Word.Document doc,
            ProvenanceAnchorInput anchor,
            int itemIndex)
        {
            string code = anchor?.Code;
            int occurrenceIndex = anchor?.Index ?? 0;
            string tableId = anchor?.TableId;

            Word.Range range = SentenceCodeLocator.LocateRange(
                doc,
                code,
                occurrenceIndex,
                tableId,
                tableScope: null,
                allowAutoTableScope: true,
                debugTag: $"provenance:{code}");

            if (range == null)
            {
                return new ProvenanceValidationResult
                {
                    Ok = false,
                    Error = $"无法定位句子 {code}",
                    FailedIndex = itemIndex,
                    AnchorCode = code,
                };
            }

            string expected = DocumentState.GetSentenceContent(code);
            if (!string.IsNullOrEmpty(expected))
            {
                string actual = NormalizeCompareText(range.Text);
                string exp = NormalizeCompareText(expected);
                if (!string.Equals(actual, exp, StringComparison.Ordinal))
                {
                    return new ProvenanceValidationResult
                    {
                        Ok = false,
                        Error = "句子内容已变更，请重新 get_document_content",
                        FailedIndex = itemIndex,
                        AnchorCode = code,
                    };
                }
            }

            return new ProvenanceValidationResult { Ok = true };
        }

        public static ProvenanceValidationResult ValidateObjectAnchor(
            Word.Document doc,
            ProvenanceAnchorInput anchor,
            ProvenanceAnchorKind kind,
            int itemIndex)
        {
            string code = anchor?.Code;
            bool exists;
            string kindLabel;
            if (kind == ProvenanceAnchorKind.Table)
            {
                exists = ObjectInsertRangeHelper.ObjectExistsAfterTable(doc, code);
                kindLabel = "表格";
            }
            else if (kind == ProvenanceAnchorKind.Chart)
            {
                exists = ObjectInsertRangeHelper.ObjectExistsAfterChart(doc, code);
                kindLabel = "图表";
            }
            else if (kind == ProvenanceAnchorKind.Image)
            {
                exists = ObjectInsertRangeHelper.ObjectExistsAfterImage(doc, code);
                kindLabel = "图片";
            }
            else
            {
                return new ProvenanceValidationResult
                {
                    Ok = false,
                    Error = $"不支持的对象锚点类型 {kind}",
                    FailedIndex = itemIndex,
                    AnchorCode = code,
                };
            }

            if (!exists)
            {
                return new ProvenanceValidationResult
                {
                    Ok = false,
                    Error = $"无法定位{kindLabel} {code}",
                    FailedIndex = itemIndex,
                    AnchorCode = code,
                };
            }

            return new ProvenanceValidationResult { Ok = true };
        }

        private static IEnumerable<string> EnumerateLines(RenderedProvenanceItem item)
        {
            if (!string.IsNullOrEmpty(item.RenderedLineForReference))
            {
                yield return item.RenderedLineForReference;
            }

            if (item.RenderedObjectLines != null)
            {
                foreach (string line in item.RenderedObjectLines)
                {
                    yield return line;
                }
            }
        }

        private static async Task<string> RenderSourceAsync(ProvenanceSourceInput src)
        {
            if (src == null)
            {
                throw new InvalidOperationException("source 不能为空");
            }

            string type = (src.Type ?? "").Trim().ToLowerInvariant();
            switch (type)
            {
                case "file":
                    return RenderFile(src);
                case "web":
                    return Require(src.Title, "title") + "，" + Require(src.Url, "url");
                case "data_fetch":
                    return "数据来源：" + Require(src.DataSource, "data_source")
                        + "，" + Require(src.Description, "description");
                case "kb":
                    return await RenderKbAsync(src).ConfigureAwait(false);
                case "text":
                    return (src.Text ?? "").Trim();
                default:
                    throw new InvalidOperationException($"未知 source.type: {src.Type}");
            }
        }

        private static string RenderFile(ProvenanceSourceInput src)
        {
            string file = Require(src.File, "file");
            if (file.Contains("/") || file.Contains("\\"))
            {
                throw new InvalidOperationException("file 须为裸文件名，不可含路径");
            }

            string path = WorkspacePathResolver.ResolveReadPath(file, ConversationContext.CurrentId);
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                throw new InvalidOperationException($"工作区文件不存在: {file}");
            }

            string display = path.Replace('\\', '/');
            return Require(src.Title, "title")
                + "，数据来源：" + Require(src.DataSource, "data_source")
                + "，文件：" + display;
        }

        private static async Task<string> RenderKbAsync(ProvenanceSourceInput src)
        {
            string uuid = Require(src.StorageDocUuid, "storage_doc_uuid");
            KbProvenanceMetaResolver.KbMeta meta = await KbProvenanceMetaResolver.ResolveAsync(uuid)
                .ConfigureAwait(false);
            return Require(src.Title, "title")
                + "，知识库：" + meta.KbName
                + "，文件：" + meta.DocumentName;
        }

        private static string Require(string value, string fieldName)
        {
            string v = (value ?? "").Trim();
            if (string.IsNullOrEmpty(v))
            {
                throw new InvalidOperationException($"缺少字段: {fieldName}");
            }

            return v;
        }

        /// <summary>
        /// 插件侧已强制按段写入；去掉模型输入/字段末尾多余换行，避免叠出空段。
        /// </summary>
        private static string NormalizeRenderedLine(string line)
        {
            if (string.IsNullOrEmpty(line))
            {
                return "";
            }

            return line.TrimEnd('\r', '\n');
        }

        private static string NormalizeCompareText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Replace("\r", "").Replace("\n", "").Trim();
        }
    }
}
