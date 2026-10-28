namespace WordAddIn1.PresentationHost
{
    internal static class PptBackgroundWppManager
    {
        public static bool TryBackground(
            object presentation,
            string channelId,
            PresentationBackgroundRequest request,
            out PresentationBackgroundResult result,
            out string error)
        {
            return PptBackgroundEngine.TryRun(
                presentation,
                channelId,
                "wpp",
                request,
                out result,
                out error);
        }
    }
}
