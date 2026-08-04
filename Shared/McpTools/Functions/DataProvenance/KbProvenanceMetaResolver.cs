using System;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1
{
    public static class KbProvenanceMetaResolver
    {
        public sealed class KbMeta
        {
            public string KbName { get; set; }
            public string DocumentName { get; set; }
        }

        public static async Task<KbMeta> ResolveAsync(string storageDocUuid)
        {
            string uuid = (storageDocUuid ?? "").Trim();
            if (string.IsNullOrEmpty(uuid))
            {
                throw new InvalidOperationException("缺少 storage_doc_uuid");
            }

            if (!UserService.Instance.CheckLoginStatus())
            {
                throw new InvalidOperationException("用户未登录，无法解析知识库来源");
            }

            string baseUrl = ConfigManager.Config?.Api?.BaseUrl;
            if (string.IsNullOrEmpty(baseUrl))
            {
                throw new InvalidOperationException("Api.BaseUrl 未配置");
            }

            string url =
                $"{baseUrl}/knowledge/document/provenance_meta?storage_doc_uuid={Uri.EscapeDataString(uuid)}";
            BackendApiClient.JsonResult result = await BackendApiClient
                .GetAuthenticatedAsync(url, retryOnUnauthorized: true, timeout: TimeSpan.FromSeconds(60))
                .ConfigureAwait(false);

            if (!result.Success)
            {
                throw new InvalidOperationException(
                    $"知识库文档解析失败（{(int)result.StatusCode}）: {uuid}");
            }

            JObject root = JObject.Parse(result.Body ?? "{}");
            if (root["success"]?.Value<bool>() != true)
            {
                string detail = root["detail"]?.ToString() ?? root["message"]?.ToString() ?? "未知错误";
                throw new InvalidOperationException(detail);
            }

            JObject data = root["data"] as JObject;
            if (data == null)
            {
                throw new InvalidOperationException("知识库文档解析响应无效");
            }

            return new KbMeta
            {
                KbName = data["kb_name"]?.ToString() ?? "",
                DocumentName = data["document_name"]?.ToString() ?? "",
            };
        }
    }
}
