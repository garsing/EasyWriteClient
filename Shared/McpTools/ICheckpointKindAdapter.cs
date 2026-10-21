namespace WordAddIn1
{
    public sealed class CheckpointLiveTarget
    {
        public ICheckpointKindAdapter Adapter { get; set; }
        public IOperationChannel Channel { get; set; }
        public string ChannelId { get; set; }
        public string DocUuid { get; set; }
        public object Live { get; set; }
    }

    public interface ICheckpointKindAdapter
    {
        string CopyExtension { get; }

        CheckpointCaret CaptureCaret(CheckpointLiveTarget target);

        void PrepareForSave(CheckpointLiveTarget target);

        void SaveCopy(CheckpointLiveTarget target, string destinationPath);

        void RestoreCopy(CheckpointLiveTarget target, string checkpointPath);

        void RestoreCaret(CheckpointLiveTarget target, CheckpointCaret caret);
    }
}
