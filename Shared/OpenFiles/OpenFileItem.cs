using Newtonsoft.Json;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// 侧栏「打开文件」列表项（bridge JSON 使用 camelCase）。
    /// </summary>
    internal sealed class OpenFileItem
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("appType")]
        public string AppType { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("fullPath")]
        public string FullPath { get; set; }

        [JsonProperty("isSaved")]
        public bool IsSaved { get; set; }
    }
}
