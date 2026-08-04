using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class F_ApplyDataProvenanceTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_data_provenance"] = async (args) =>
            {
                try
                {
                    if (wordApplication == null)
                    {
                        return Fail("Word应用程序不可用");
                    }

                    dynamic wordApp = wordApplication;
                    if (wordApp.ActiveDocument == null)
                    {
                        return Fail("没有活动的Word文档");
                    }

                    Word.Document doc = wordApp.ActiveDocument;
                    DocumentState.BindAndActivate(doc);

                    if (!DataProvenanceJsonParser.TryParseItems(args, out List<ProvenanceItemInput> items, out string parseError))
                    {
                        return Fail(parseError);
                    }

                    ProvenanceValidationResult structure = DataProvenanceParamValidator.ValidateStructure(items);
                    if (!structure.Ok)
                    {
                        return Fail(structure.Error);
                    }

                    // 不读写 TrackRevisions（2026-08：取消审阅工具注册与默认关审阅）
                    WordDocumentExtractor.ProcessDocument(
                        doc,
                        ProcessDocumentOptions.ForProcessActions("apply_data_provenance"));

                    var rendered = new List<RenderedProvenanceItem>();
                    for (int i = 0; i < items.Count; i++)
                    {
                        AgentRunCancellation.ThrowIfCancelled();
                        ProvenanceItemInput item = items[i];
                        ProvenanceAnchorKind kind = DataProvenanceParamValidator.ResolveKind(item.Anchor?.Code);

                        ProvenanceValidationResult anchorCheck = kind == ProvenanceAnchorKind.Sentence
                            ? DataProvenanceSourceRenderer.ValidateSentenceAnchor(doc, item.Anchor, i)
                            : DataProvenanceSourceRenderer.ValidateObjectAnchor(doc, item.Anchor, kind, i);
                        if (!anchorCheck.Ok)
                        {
                            return FailAt(anchorCheck);
                        }

                        (RenderedProvenanceItem renderedItem, string renderError) =
                            await DataProvenanceSourceRenderer.RenderItemAsync(item, i).ConfigureAwait(false);
                        if (renderError != null)
                        {
                            return FailAt(i, item.Anchor?.Code, renderError);
                        }

                        renderedItem.Kind = kind;
                        renderedItem.Anchor = item.Anchor;
                        rendered.Add(renderedItem);
                    }

                    List<string> warnings = DataProvenanceClearHelper.ClearIfNeeded(doc);

                    int sentenceNo = 0;
                    var sentenceItems = new List<RenderedProvenanceItem>();
                    int tableCount = 0;
                    int chartCount = 0;
                    int imageCount = 0;

                    foreach (RenderedProvenanceItem item in rendered)
                    {
                        if (item.Kind == ProvenanceAnchorKind.Sentence)
                        {
                            continue;
                        }

                        DataProvenanceWriteHelper.WriteObjectLines(doc, item);
                        if (item.Kind == ProvenanceAnchorKind.Table)
                        {
                            tableCount++;
                        }
                        else if (item.Kind == ProvenanceAnchorKind.Chart)
                        {
                            chartCount++;
                        }
                        else if (item.Kind == ProvenanceAnchorKind.Image)
                        {
                            imageCount++;
                        }
                    }

                    foreach (RenderedProvenanceItem item in rendered)
                    {
                        if (item.Kind != ProvenanceAnchorKind.Sentence)
                        {
                            continue;
                        }

                        sentenceNo++;
                        item.SentenceNumber = sentenceNo;
                        sentenceItems.Add(item);
                    }

                    List<RenderedProvenanceItem> markerOrder =
                        DataProvenanceWriteHelper.SortSentenceItemsForMarkerWrite(doc, sentenceItems);
                    foreach (RenderedProvenanceItem item in markerOrder)
                    {
                        DataProvenanceWriteHelper.WriteSentenceMarker(
                            doc,
                            item,
                            item.SentenceNumber.Value);
                    }

                    DataProvenanceWriteHelper.WriteReferenceSection(doc, sentenceItems);

                    WordDocumentExtractor.ProcessDocument(
                        doc,
                        ProcessDocumentOptions.ForProcessActions("apply_data_provenance_post"));

                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            sentence_count = sentenceNo,
                            table_count = tableCount,
                            chart_count = chartCount,
                            image_count = imageCount,
                            item_count = items.Count,
                            warnings,
                        },
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ApplyDataProvenance] 执行失败: {ex.Message}");
                    return Fail($"执行失败: {ex.Message}");
                }
            };
        }

        private static ToolResult Fail(string error) =>
            new ToolResult { Success = false, Error = error };

        private static ToolResult FailAt(ProvenanceValidationResult validation) =>
            new ToolResult
            {
                Success = false,
                Error = validation.Error,
                Data = new
                {
                    failed_index = validation.FailedIndex,
                    anchor_code = validation.AnchorCode,
                },
            };

        private static ToolResult FailAt(int index, string anchorCode, string error) =>
            new ToolResult
            {
                Success = false,
                Error = error,
                Data = new
                {
                    failed_index = index,
                    anchor_code = anchorCode,
                },
            };
    }
}
