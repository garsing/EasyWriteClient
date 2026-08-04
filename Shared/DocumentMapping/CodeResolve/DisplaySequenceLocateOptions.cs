namespace WordAddIn1.DocumentMapping.CodeResolve
{
    public enum CodeStreamKind
    {
        SentenceReadText = 0,
        ParagraphDisplay = 1
    }

    public sealed class DisplaySequenceLocateOptions
    {
        public string TableId { get; set; }

        public CodeStreamKind StreamKind { get; set; } = CodeStreamKind.SentenceReadText;

        public string Snapshot { get; set; }

        public string SegSnapshot { get; set; }

        public TableScopeIndex TableScope { get; set; }
    }
}
