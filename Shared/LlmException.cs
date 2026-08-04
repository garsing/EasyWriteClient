using System;

namespace WordAddIn1
{
    /// <summary>
    /// LLM / 认证相关异常（供 Auth、LlmClient 等共用）。
    /// </summary>
    public class LlmException : Exception
    {
        public LlmException(string message) : base(message)
        {
        }

        public LlmException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
