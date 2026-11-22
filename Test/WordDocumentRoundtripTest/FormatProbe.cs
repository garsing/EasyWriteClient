using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    internal sealed class Marker
    {
        public string Code;
        public string Text;
    }

    /// <summary>临时稿、读码、搜格式、读层。用完关自己的稿。</summary>
    internal sealed class FormatProbe : IDisposable
    {
        private static readonly Regex SentencePattern = new Regex(
            @"\[(S_[0-9A-Za-z]+)-start\]([\s\S]*?)\[\1-end\]",
            RegexOptions.Compiled);

        private static readonly Regex ParagraphPattern = new Regex(
            @"\[(P_[0-9A-Za-z]+(?:_[1-9][0-9]*)?)(?:-start| start)\]([\s\S]*?)\[\1(?:-end| end)\]",
            RegexOptions.Compiled);

        public WdSession Session;
        public string Path;
        public string Id;
        public bool Ready;
        public string SessionDir;
        public List<Marker> Sentences = new List<Marker>();
        public List<Marker> Paragraphs = new List<Marker>();

        public static FormatProbe Open(WdSession session, string text, string sessionDir)
        {
            var probe = new FormatProbe
            {
                Session = session,
                SessionDir = sessionDir,
                Path = session.NewDocxPath()
            };
            ToolResult opened = session.Open(probe.Path, true, session.App);
            if (opened == null || !opened.Success)
            {
                probe.Id = opened == null ? "打开失败" : (opened.Error ?? "打开失败");
                return probe;
            }

            probe.Id = WdSession.Str(opened, "channel_id");
            probe.Ready = !string.IsNullOrEmpty(probe.Id) && session.TypePlain(probe.Path, text ?? "");
            if (!probe.Ready && string.IsNullOrEmpty(probe.Id))
            {
                probe.Id = "铺正文失败";
            }

            return probe;
        }

        public string Snapshot()
        {
            ToolResult got = Session.GetContent(Id);
            return got != null && got.Success ? WdSession.Str(got, "document_content") : "";
        }

        public bool ReadSentences()
        {
            ToolResult got = Session.GetContent(Id);
            Sentences = Parse(got != null && got.Success ? WdSession.Str(got, "document_content") : "", SentencePattern);
            return Sentences.Count > 0;
        }

        public bool ReadParagraphs()
        {
            ToolResult got = Session.GetContent(Id, "paragraph");
            Paragraphs = Parse(got != null && got.Success ? WdSession.Str(got, "document_content") : "", ParagraphPattern);
            return Paragraphs.Count > 0;
        }

        public string Sentence(string needle)
        {
            return Nth(Sentences, needle, 0);
        }

        public string Paragraph(string needle)
        {
            return Nth(Paragraphs, needle, 0);
        }

        public bool Touch(Func<object, bool> edit)
        {
            if (edit == null)
            {
                return false;
            }

            return Session.IsWps
                ? Session.TryWpsDoc(Path, edit)
                : Session.OnWordDoc(Path, doc => edit(doc));
        }

        public bool EditParagraph(string needle, Action<object> edit)
        {
            return Session.EditParagraph(Path, needle, edit);
        }

        public void Dispose()
        {
            if (Session != null && !string.IsNullOrEmpty(Path))
            {
                Session.CloseDocByPath(Path);
            }
        }

        public static JObject Data(ToolResult result)
        {
            if (result == null || result.Data == null)
            {
                return new JObject();
            }

            try
            {
                return JObject.Parse(JsonConvert.SerializeObject(result.Data));
            }
            catch
            {
                return new JObject();
            }
        }

        public static string FileOf(string sessionDir, string name)
        {
            return string.IsNullOrEmpty(sessionDir) ? "" : System.IO.Path.Combine(sessionDir, name);
        }

        public static void DeleteSearchFiles(string sessionDir)
        {
            foreach (string name in new[] { "format_search_character.json", "format_search_paragraph.json" })
            {
                string full = FileOf(sessionDir, name);
                if (!string.IsNullOrEmpty(full) && File.Exists(full))
                {
                    File.Delete(full);
                }
            }
        }

        private static List<Marker> Parse(string display, Regex pattern)
        {
            var list = new List<Marker>();
            if (string.IsNullOrEmpty(display))
            {
                return list;
            }

            foreach (Match match in pattern.Matches(display))
            {
                list.Add(new Marker { Code = match.Groups[1].Value, Text = match.Groups[2].Value });
            }

            return list;
        }

        private static string Nth(List<Marker> markers, string needle, int nth)
        {
            int seen = 0;
            foreach (Marker marker in markers)
            {
                if ((marker.Text ?? "").IndexOf(needle, StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                if (seen == nth)
                {
                    return marker.Code;
                }

                seen++;
            }

            return "";
        }
    }
}
