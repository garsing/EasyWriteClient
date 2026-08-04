using System.Collections.Generic;

namespace WordAddIn1
{
    public enum ProvenanceAnchorKind
    {
        Sentence,
        Table,
        Chart,
        Image
    }

    public sealed class ProvenanceAnchorInput
    {
        public string Code { get; set; }
        public int? Index { get; set; }
        public string TableId { get; set; }
    }

    public sealed class ProvenanceSourceInput
    {
        public string Type { get; set; }
        public string File { get; set; }
        public string DataSource { get; set; }
        public string Title { get; set; }
        public string Url { get; set; }
        public string Description { get; set; }
        public string StorageDocUuid { get; set; }
        public string Text { get; set; }
    }

    public sealed class ProvenanceItemInput
    {
        public ProvenanceAnchorInput Anchor { get; set; }
        public ProvenanceSourceInput Source { get; set; }
        public List<ProvenanceSourceInput> Sources { get; set; }
    }

    public sealed class RenderedProvenanceItem
    {
        public ProvenanceAnchorKind Kind { get; set; }
        public ProvenanceAnchorInput Anchor { get; set; }
        public int? SentenceNumber { get; set; }
        public string RenderedLineForReference { get; set; }
        public List<string> RenderedObjectLines { get; set; } = new List<string>();
    }

    public sealed class ProvenanceValidationResult
    {
        public bool Ok { get; set; }
        public string Error { get; set; }
        public int? FailedIndex { get; set; }
        public string AnchorCode { get; set; }
    }
}
