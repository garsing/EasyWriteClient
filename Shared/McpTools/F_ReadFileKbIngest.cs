using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// F_read_file detailed=true：上传 default_upload、等 process、取 model_view。
    /// </summary>
    internal static class F_ReadFileKbIngest
    {
        public const int TimeoutSeconds = 600;

        public static async Task<ToolResult> RunAsync(string localPath)
        {
            HostChatUploadResult upload = await KnowledgeBaseService
                .UploadDocumentForChatFromPathAsync(localPath)
                .ConfigureAwait(false);
            if (upload == null || !upload.Success || string.IsNullOrEmpty(upload.StorageDocUuid))
            {
                return new ToolResult
                {
                    Success = false,
                    Error = string.IsNullOrEmpty(upload?.Message) ? "上传失败" : upload.Message
                };
            }

            bool deduplicated = upload.SkipProcess;
            if (!upload.SkipProcess)
            {
                string processError = await KnowledgeBaseService
                    .ProcessDocumentForReadAsync(upload.StorageDocUuid, TimeoutSeconds)
                    .ConfigureAwait(false);
                if (!string.IsNullOrEmpty(processError))
                {
                    return new ToolResult { Success = false, Error = processError };
                }
            }

            JObject view;
            try
            {
                view = await KnowledgeBaseService
                    .GetDocumentModelViewAsync(upload.StorageDocUuid, TimeoutSeconds)
                    .ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                return new ToolResult { Success = false, Error = ex.Message };
            }

            if (view == null)
            {
                return new ToolResult { Success = false, Error = "获取文档视图失败" };
            }

            if (deduplicated)
            {
                view["deduplicated"] = true;
            }

            return new ToolResult
            {
                Success = true,
                Data = view.ToObject<Dictionary<string, object>>() ?? new Dictionary<string, object>()
            };
        }
    }
}
