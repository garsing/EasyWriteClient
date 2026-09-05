using System;
using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    public sealed class PresentationManageSlideRequest
    {
        public string Action { get; set; }

        public string SlideId { get; set; }

        /// <summary>1-based；null = add/duplicate 目标稿尾部。</summary>
        public int? ToIndex { get; set; }

        public string Layout { get; set; }

        public string[] Order { get; set; }

        public bool Confirm { get; set; }

        public string[] SlideIds { get; set; }

        public string SourceChannelId { get; set; }

        public bool TryResolveSlideSequence(out string[] sequence, out string error)
        {
            sequence = null;
            error = null;
            if (SlideIds != null)
            {
                if (SlideIds.Length == 0)
                {
                    error = "slide_ids 不能为空";
                    return false;
                }

                var items = new List<string>(SlideIds.Length);
                for (int i = 0; i < SlideIds.Length; i++)
                {
                    string id = (SlideIds[i] ?? "").Trim();
                    if (string.IsNullOrEmpty(id))
                    {
                        error = "slide_ids 含空 slide_id";
                        return false;
                    }

                    items.Add(id);
                }

                sequence = items.ToArray();
                return true;
            }

            if (!string.IsNullOrWhiteSpace(SlideId))
            {
                sequence = new[] { SlideId.Trim() };
                return true;
            }

            error = "须提供 slide_id 或 slide_ids";
            return false;
        }

        public static bool SequenceHasDuplicate(string[] sequence, out string duplicated)
        {
            duplicated = null;
            if (sequence == null)
            {
                return false;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string id in sequence)
            {
                if (!seen.Add(id))
                {
                    duplicated = id;
                    return true;
                }
            }

            return false;
        }
    }

    public sealed class CreatedSlideInfo
    {
        public string SourceSlideId { get; set; }

        public string SlideId { get; set; }

        public int Index { get; set; }
    }

    public sealed class PresentationManageSlideResult
    {
        public string ChannelId { get; set; }

        public string SourceChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        public string FocusSlideId { get; set; }

        public int? FocusIndex { get; set; }

        public int SlideCount { get; set; }

        public List<PresentationSlideInfo> Slides { get; set; }

        public List<ClearedPlaceholderInfo> ClearedPlaceholders { get; set; }

        public List<CreatedSlideInfo> Created { get; set; }

        public List<string> Deleted { get; set; }
    }
}
