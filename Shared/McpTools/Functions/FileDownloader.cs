using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 文件下载器类 - 从云端下载文件（实现见 <see cref="BackendApiClient"/>）。
    /// </summary>
    public static class FileDownloader
    {
        /// <summary>
        /// 从用户工作区下载文件到本地（relative_path 如 sessions/{id}/{filename}）。
        /// </summary>
        public static Task<bool> DownloadFileFromUserDirectory(string relativePath, string localPath)
        {
            return BackendApiClient.DownloadFileFromUserDirectoryAsync(relativePath, localPath);
        }
    }
}
