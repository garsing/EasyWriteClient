using System;
using System.Collections.Generic;
using WordAddIn1.OpenFiles;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 一次工具调用读一份文档默认。字符是 docDefaults 的 rPr，段落是 docDefaults 的 pPr。
    /// </summary>
    public sealed class FormatDefaultSource : IDisposable
    {
        [ThreadStatic]
        private static FormatDefaultSource current;

        [ThreadStatic]
        private static int depth;

        private FormatDefaultSource previous;

        public Dictionary<string, object> Character { get; private set; }

        public Dictionary<string, object> Paragraph { get; private set; }

        public static FormatDefaultSource Current => depth > 0 ? current : null;

        public static FormatDefaultSource Begin(Word.Document document)
        {
            var scope = new FormatDefaultSource();
            scope.previous = current;
            if (depth == 0)
            {
                FormatDefaultSource loaded = Load(document);
                scope.Character = loaded.Character;
                scope.Paragraph = loaded.Paragraph;
                current = scope;
            }
            else
            {
                scope.Character = current?.Character;
                scope.Paragraph = current?.Paragraph;
                current = scope;
            }

            depth++;
            return scope;
        }

        public void Dispose()
        {
            if (depth > 0)
            {
                depth--;
            }

            current = previous;
        }

        public static FormatDefaultSource Load(Word.Document document)
        {
            var source = new FormatDefaultSource
            {
                Character = new Dictionary<string, object>(StringComparer.Ordinal),
                Paragraph = new Dictionary<string, object>(StringComparer.Ordinal),
            };
            if (document == null)
            {
                return source;
            }

            string raw = ReadFullXml(document);
            if (string.IsNullOrEmpty(raw))
            {
                return source;
            }

            OpenXmlPackage package;
            try
            {
                package = OpenXmlPackage.Parse(raw);
            }
            catch
            {
                return source;
            }

            try
            {
                Dictionary<string, object> character = CharFormatRunReader.ReadDocCharDefaults(package);
                if (character != null)
                {
                    source.Character = character;
                }
            }
            catch
            {
            }

            try
            {
                Dictionary<string, object> paragraph = ParaFormatOoxmlReader.LoadDocDefaults(package.Xml);
                if (paragraph != null)
                {
                    source.Paragraph = paragraph;
                }
            }
            catch
            {
            }

            return source;
        }

        private static string ReadFullXml(Word.Document document)
        {
            try
            {
                string raw = document.WordOpenXML;
                if (!string.IsNullOrEmpty(raw))
                {
                    return raw;
                }
            }
            catch
            {
            }

            try
            {
                string late = Convert.ToString(WpsCom.GetProperty(document, "WordOpenXML") ?? "");
                if (!string.IsNullOrEmpty(late))
                {
                    return late;
                }
            }
            catch
            {
            }

            return null;
        }
    }
}
