using System.Collections.Generic;

namespace WordAddIn1
{
    public enum OperationLogEntryKind
    {
        Checkpoint,
        Action
    }

    public sealed class OperationLogEntry
    {
        public OperationLogEntryKind Kind { get; set; }
        /// <summary>副本条目：timestamp，不含「副本」前缀。</summary>
        public string Timestamp { get; set; }
        public string ToolName { get; set; }
        /// <summary>action 行 JSON 部分（工具名之后）。</summary>
        public string ActionLineBody { get; set; }
    }

    public sealed class DocumentCheckpointSession
    {
        public string DocUuid { get; set; }
        public List<OperationLogEntry> Entries { get; } = new List<OperationLogEntry>();
        public List<string> CheckpointTimestampsInOrder { get; } = new List<string>();
    }

    public sealed class CheckpointBeforeResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string Timestamp { get; set; }
        public string DocUuid { get; set; }
    }
}
