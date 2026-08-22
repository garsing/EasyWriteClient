using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordAddIn1
{
    public static class F_WorkspaceInfoTool
    {
        public const string Hint =
            "区内其它工具的 path 只写相对名或裸文件名，不要把 session_dir 拼进 F_write_file / F_read_file。";

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_workspace_info"] = args =>
            {
                try
                {
                    if (!UserService.Instance.CheckLoginStatus()
                        || string.IsNullOrWhiteSpace(UserService.Instance.WorkspaceRootEffective))
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = "用户未登录或工作区未初始化"
                        });
                    }

                    string sessionDir = WorkspacePathResolver.GetSessionDirectory();
                    string sessionNorm = NormalizeDir(sessionDir, trailingSlash: true);
                    string rootNorm = NormalizeDir(UserService.Instance.WorkspaceRootEffective, trailingSlash: false);
                    string conversationId = WorkspaceReconcile.NormalizeConversationId(ConversationContext.CurrentId);

                    return Task.FromResult(new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            session_dir = sessionNorm,
                            workspace_root = rootNorm,
                            conversation_id = conversationId,
                            hint = Hint
                        }
                    });
                }
                catch (Exception ex)
                {
                    return Task.FromResult(new ToolResult
                    {
                        Success = false,
                        Error = "获取工作区信息失败: " + ex.Message
                    });
                }
            };
        }

        private static string NormalizeDir(string path, bool trailingSlash)
        {
            string norm = (path ?? "").Replace('\\', '/').TrimEnd('/');
            if (string.IsNullOrEmpty(norm))
            {
                return norm;
            }

            return trailingSlash ? norm + "/" : norm;
        }
    }
}
