using System;
using System.Collections.Generic;

namespace WordAddIn1
{
    /// <summary>
    /// 单个 Word 文档实例在插件会话内的状态分片（按 doc_uuid 隔离）。
    /// </summary>
    public sealed class DocumentSessionState
    {
        public string DocUuid { get; set; } = "";

        /// <summary>[句子名称, 哈希码, 内容, 标记]</summary>
        public List<List<string>> SentenceNameMapping { get; } = new List<List<string>>();

        /// <summary>[表格编号, 哈希码, 表格内容]（文档内累积）</summary>
        public List<List<string>> TableIdMapping { get; } = new List<List<string>>();

        public List<string> TableIdOrder { get; } = new List<string>();

        /// <summary>[图表编号, 哈希码, 图表体内 XML]（文档内累积）</summary>
        public List<List<string>> ChartIdMapping { get; } = new List<List<string>>();

        public List<string> ChartIdOrder { get; } = new List<string>();

        /// <summary>[图片编号 I_, 最终材料哈希, meta 原文, salt_k]（按实例唯一，每次读文重建）</summary>
        public List<List<string>> ImageIdMapping { get; } = new List<List<string>>();

        public List<string> ImageIdOrder { get; } = new List<string>();

        public int NextNameIndex { get; set; }

        public int NextParagraphNameIndex { get; set; }

        /// <summary>[段落名称 P_, 哈希码, storedText（含 \r 规则）, segMark]</summary>
        public List<List<string>> ParagraphNameMapping { get; } = new List<List<string>>();

        public Dictionary<string, string> SentenceToParagraph { get; } =
            new Dictionary<string, string>(StringComparer.Ordinal);

        public Dictionary<string, List<string>> ParagraphToSentences { get; } =
            new Dictionary<string, List<string>>(StringComparer.Ordinal);

        /// <summary>按 chunk 序存的 paragraph display（仅 [P_*]）</summary>
        public List<string> ChunkParagraphDisplayContents { get; } = new List<string>();

        public string Snapshot { get; set; } = "";

        public string SegSnapshot { get; set; } = "";

        public List<SnapshotRecord> SnapshotHistory { get; } = new List<SnapshotRecord>();

        public int NextSnapshotHistoryIndex { get; set; }

        public ProcessDocumentCacheEntry LastProcessCache { get; set; }

        public Dictionary<string, Dictionary<string, object>> PreReplaceByNewCode { get; } =
            new Dictionary<string, Dictionary<string, object>>(StringComparer.Ordinal);

        public int SnapshotIdx { get; set; } = -1;

        public Dictionary<string, Dictionary<string, object>> SentenceNameToHeadingMap { get; } =
            new Dictionary<string, Dictionary<string, object>>();

        public List<Dictionary<string, object>> SubtypeFormatMapping { get; } =
            new List<Dictionary<string, object>>();

        public List<Dictionary<string, object>> RecognizedHeadingsInDocumentOrder { get; } =
            new List<Dictionary<string, object>>();

        /// <summary>正文轻量抽取众数；与 SubtypeFormatMapping 隔离。</summary>
        public Dictionary<string, object> BodyFormat { get; set; }

        /// <summary>抽格式时记下的源 body 指纹；四类产物共用一份。</summary>
        public string FormatExtractBodyFingerprint { get; set; }

        public Dictionary<string, object> SectionPageLayout { get; set; }

        public List<Dictionary<string, object>> TableFormatCatalog { get; set; }

        public Dictionary<string, object> TableFormatStandard { get; set; }

        public Dictionary<string, object> ParaFormatClusters { get; set; }

        public void ClearFormatExtractProducts()
        {
            SubtypeFormatMapping.Clear();
            SectionPageLayout = null;
            TableFormatCatalog = null;
            TableFormatStandard = null;
            ParaFormatClusters = null;
        }

        /// <summary>UUID 变更后清除易失状态；checkpoint restore 后亦调用（I6/B：不清 SentenceNameMapping）。</summary>
        public void ClearVolatileState()
        {
            PreReplaceByNewCode.Clear();
            LastProcessCache = null;
            SnapshotHistory.Clear();
            NextSnapshotHistoryIndex = 0;
            Snapshot = "";
            SegSnapshot = "";
            SnapshotIdx = -1;
            ChunkParagraphDisplayContents.Clear();
            BodyFormat = null;
        }
    }
}
