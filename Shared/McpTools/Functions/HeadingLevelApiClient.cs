using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// 调用服务端 POST /heading_level/infer，获取各 detailed_subtype 的大纲层级（level 越小越顶层）。
    /// </summary>
    public static class HeadingLevelApiClient
    {
        /// <summary>
        /// 同步调用；失败或空输入时返回空字典，不抛异常。
        /// </summary>
        public static Dictionary<string, int> TryInferHeadingLevels(string headingsBlock, List<string> uniqueTypes)
        {
            if (string.IsNullOrWhiteSpace(headingsBlock) || uniqueTypes == null || uniqueTypes.Count == 0)
            {
                return new Dictionary<string, int>();
            }

            try
            {
                return InferAsync(headingsBlock, uniqueTypes).ConfigureAwait(false).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HeadingLevelApiClient] 标题层级推断异常: {ex.Message}");
                return new Dictionary<string, int>();
            }
        }

        private static async Task<Dictionary<string, int>> InferAsync(string headingsBlock, List<string> uniqueTypes)
        {
            var userService = UserService.Instance;
            if (!userService.CheckLoginStatus() || string.IsNullOrEmpty(userService.AccessToken))
            {
                System.Diagnostics.Debug.WriteLine("[HeadingLevelApiClient] 未登录或缺少 token，跳过推断");
                return new Dictionary<string, int>();
            }

            string baseUrl = (ConfigManager.Config.Api.BaseUrl ?? "").TrimEnd('/');
            if (string.IsNullOrEmpty(baseUrl))
            {
                System.Diagnostics.Debug.WriteLine("[HeadingLevelApiClient] Api.BaseUrl 为空，跳过推断");
                return new Dictionary<string, int>();
            }

            string url = $"{baseUrl}/heading_level/infer";

            var body = new JObject
            {
                ["headings_block"] = headingsBlock,
                ["unique_types"] = JArray.FromObject(uniqueTypes)
            };

            System.Diagnostics.Debug.WriteLine($"[HeadingLevelApiClient] POST {url}");

            var postResult = await BackendApiClient.PostJsonAuthenticatedAsync(
                url,
                body.ToString(),
                retryOnUnauthorized: true,
                TimeSpan.FromSeconds(120)).ConfigureAwait(false);

            if (!postResult.Success)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[HeadingLevelApiClient] HTTP {(int)postResult.StatusCode}: {postResult.Body}");
                return new Dictionary<string, int>();
            }

            string responseText = postResult.Body;
            var root = JObject.Parse(responseText);
            if (root["success"]?.Value<bool>() != true || root["data"] == null)
            {
                System.Diagnostics.Debug.WriteLine($"[HeadingLevelApiClient] 响应 success=false 或无 data: {responseText}");
                return new Dictionary<string, int>();
            }

            var levelsToken = root["data"]?["levels"] as JArray;
            if (levelsToken == null || levelsToken.Count == 0)
            {
                return new Dictionary<string, int>();
            }

            var map = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var el in levelsToken)
            {
                string dst = el["detailed_subtype"]?.ToString();
                if (string.IsNullOrEmpty(dst))
                {
                    continue;
                }

                int level = 0;
                JToken lv = el["level"];
                if (lv != null && lv.Type != JTokenType.Null)
                {
                    if (lv.Type == JTokenType.Integer)
                    {
                        level = lv.Value<int>();
                    }
                    else
                    {
                        int.TryParse(lv.ToString(), out level);
                    }
                }

                if (level > 0)
                {
                    map[dst] = level;
                }
            }

            System.Diagnostics.Debug.WriteLine($"[HeadingLevelApiClient] 已解析 {map.Count} 条 detailed_subtype 层级");
            return map;
        }
    }
}
