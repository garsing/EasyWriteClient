using System;
using System.Collections.Generic;
using WordAddIn1.PresentationHost;

namespace PptHtmlContentRoundtripTest
{
    internal sealed class ContentCase
    {
        public int Index { get; set; }
        public int Batch { get; set; } = 1;
        public int Page { get; set; }
        public string Name { get; set; }
        public Func<ContentAssets, string> CreateHtml { get; set; }
        public Func<ContentAssets, IList<string>, string> ReplaceHtml { get; set; }
        public string ExpectApplyErrorContains { get; set; }
        public Func<ContentAssertContext, string> Match { get; set; }
        public bool ContentOnly { get; set; }
    }

    internal sealed class ContentAssets
    {
        public string RedPng { get; set; }
        public string BluePng { get; set; }
    }

    internal sealed class ContentAssertContext
    {
        public ContentCase Case { get; set; }
        public IList<string> CreatedShapeIds { get; set; }
        public PptHtmlReadResult AfterCreate { get; set; }
        public PptHtmlReadResult AfterReplace { get; set; }
        public PptHtmlApplyResult CreateResult { get; set; }
        public PptHtmlApplyResult ReplaceResult { get; set; }
        public string ExportDir { get; set; }
        public object Presentation { get; set; }
    }
}
