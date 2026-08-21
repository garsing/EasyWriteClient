using System;
using System.Collections.Generic;

namespace WordAddIn1.Terminal
{
    internal static class TerminalSessionHub
    {
        private static readonly object Gate = new object();
        private static readonly Dictionary<string, TerminalSession> Sessions =
            new Dictionary<string, TerminalSession>(StringComparer.Ordinal);

        public static TerminalSession GetOrCreate(string conversationId, string cwd)
        {
            string key = Normalize(conversationId);
            lock (Gate)
            {
                if (Sessions.TryGetValue(key, out TerminalSession existing)
                    && existing != null)
                {
                    return existing;
                }

                var created = TerminalSession.Start(cwd);
                Sessions[key] = created;
                return created;
            }
        }

        public static bool TryGet(string conversationId, out TerminalSession session)
        {
            string key = Normalize(conversationId);
            lock (Gate)
            {
                return Sessions.TryGetValue(key, out session) && session != null;
            }
        }

        public static bool Dispose(string conversationId)
        {
            string key = Normalize(conversationId);
            TerminalSession session;
            lock (Gate)
            {
                if (!Sessions.TryGetValue(key, out session))
                {
                    return false;
                }

                Sessions.Remove(key);
            }

            session?.Dispose();
            return session != null;
        }

        private static string Normalize(string conversationId)
        {
            return string.IsNullOrWhiteSpace(conversationId) ? "-1" : conversationId.Trim();
        }
    }
}
