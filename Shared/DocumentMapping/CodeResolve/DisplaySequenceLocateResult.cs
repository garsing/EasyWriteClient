using System.Collections.Generic;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    public sealed class DisplaySequenceLocateResult
    {
        public int MatchCount { get; set; }

        public int? DisplayStartPosition { get; set; }

        public string ErrorCode { get; set; }

        public List<int> MatchStartPositions { get; set; } = new List<int>();

        public IReadOnlyList<string> OrderedDomainCodes { get; set; }

        public int[] SegIndexByPosition { get; set; }
    }
}
