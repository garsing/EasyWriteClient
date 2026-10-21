namespace WordAddIn1
{
    /// <summary>建副本时的光标，门面不解读字段，由类型适配还原。</summary>
    public sealed class CheckpointCaret
    {
        public int SelectionStart { get; set; }
        public int SelectionEnd { get; set; }
        public string SheetName { get; set; }
        public string CellAddress { get; set; }
        public int SlideIndex { get; set; }
        public int ShapeId { get; set; }
        public bool Captured { get; set; }
    }
}
