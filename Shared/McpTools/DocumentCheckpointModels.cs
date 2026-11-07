using System;
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
        public int Branch { get; set; } = 1;
        public bool UserRollback { get; set; }
        public string Hash { get; set; }
        public bool Retired { get; set; }
        public List<CheckpointOperationRecord> Operations { get; } = new List<CheckpointOperationRecord>();
    }

    public sealed class CheckpointOperationRecord
    {
        public string Tool { get; set; }
        public string Arguments { get; set; }
    }

    public sealed class CheckpointBranch
    {
        public int Id { get; set; }
        public int ForkedFromBranch { get; set; }
        public string ForkedFromTimestamp { get; set; }
    }

    public sealed class DocumentCheckpointSession
    {
        public string DocUuid { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public string Kind { get; set; }
        public int CurrentBranch { get; set; } = 1;
        public int PositionBranch { get; set; }
        public string PositionTimestamp { get; set; }
        public List<CheckpointBranch> Branches { get; } = new List<CheckpointBranch>();
        public List<OperationLogEntry> Entries { get; } = new List<OperationLogEntry>();
        public List<string> CheckpointTimestampsInOrder { get; } = new List<string>();
        public Dictionary<string, CheckpointCaret> CaretByTimestamp { get; } =
            new Dictionary<string, CheckpointCaret>(StringComparer.Ordinal);
    }

    public sealed class CheckpointBeforeResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string Timestamp { get; set; }
        public string DocUuid { get; set; }
        /// <summary>创建副本时的渠道号；回滚按它找活文档，不依赖 DocumentIdentity 与渠道 uuid 是否同一套。</summary>
        public string ChannelId { get; set; }
    }
}
