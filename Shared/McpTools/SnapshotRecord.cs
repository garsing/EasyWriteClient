using System;

namespace WordAddIn1
{
    /// <summary>
    /// 单次 ProcessDocument / SetSnapshots 产生的快照记录（仅句子编号，不含全文）。
    /// </summary>
    public sealed class SnapshotRecord
    {
        /// <summary>会话内单调递增，从 0 开始。</summary>
        public int Index { get; set; }

        public DateTime CreatedAtUtc { get; set; }

        /// <summary>扁平句子码列表，逗号分隔。</summary>
        public string Snapshot { get; set; }

        /// <summary>按 seg 分组的句子码，分号分隔 seg，seg 内逗号分隔。</summary>
        public string SegSnapshot { get; set; }

        /// <summary>触发来源，如 process_actions_step1、get_document_content。</summary>
        public string Source { get; set; }

        /// <summary>记录写入时的 DocumentState.CurrDocUuid。</summary>
        public string DocUuid { get; set; }
    }
}
