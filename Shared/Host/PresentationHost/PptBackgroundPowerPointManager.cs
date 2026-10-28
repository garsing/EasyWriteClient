namespace WordAddIn1.PresentationHost
{
    internal static class PptBackgroundPowerPointManager
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
                "ppt",
                request,
                out result,
                out error);
        }
    }
}
