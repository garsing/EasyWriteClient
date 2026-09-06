using System.Collections.Generic;

namespace WordAddIn1.HostPlatform
{
    public sealed class OpenDocumentResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string DocUuid { get; set; }

        public string Path { get; set; }

        public string Name { get; set; }

        public bool Created { get; set; }

        public bool Reused { get; set; }

        public bool IndexReady { get; set; }

        public string IndexError { get; set; }

        public ToolResult ToToolResult()
        {
            var data = new Dictionary<string, object>
            {
                ["channel_id"] = ChannelRegistry.ToPublicId(ChannelId) ?? "",
                ["kind"] = Kind ?? "",
                ["path"] = Path ?? "",
                ["name"] = Name ?? "",
                ["created"] = Created,
                ["reused"] = Reused,
                ["index_ready"] = IndexReady,
            };
            if (!string.IsNullOrEmpty(IndexError))
            {
                data["index_error"] = IndexError;
            }

            return new ToolResult { Success = true, Data = data };
        }
    }
}
