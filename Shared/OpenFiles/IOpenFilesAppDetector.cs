using System;
using System.Collections.Generic;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// 按应用类型探测本机已打开文档（快照 + 开/关事件）。
    /// </summary>
    internal interface IOpenFilesAppDetector : IDisposable
    {
        string AppType { get; }

        /// <summary>附着已运行实例；禁止 create 新进程。</summary>
        bool TryAttach();

        IReadOnlyList<OpenFileItem> Snapshot();

        event Action<OpenFileItem> DocumentOpened;

        event Action<string> DocumentClosed;

        /// <summary>COM 失效时触发，Monitor 清空该 app 条目并等待重附着。</summary>
        event Action Detached;

        bool IsAttached { get; }
    }
}
