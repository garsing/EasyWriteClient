using System;
using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    /// <summary>I5：add 后清掉仍为空的标题/正文/副标题占位符。</summary>
    internal static class PptEmptyPlaceholderClear
    {
        public const int Title = 1;
        public const int Body = 2;
        public const int CenterTitle = 3;
        public const int Subtitle = 4;
        public const int VerticalTitle = 16;
        public const int VerticalBody = 17;

        private static readonly string[] EmptyPrompts =
        {
            "单击此处添加标题",
            "单击此处添加文本",
            "单击此处添加副标题",
            "click to add title",
            "click to add text",
            "click to add subtitle",
            "此处添加标题",
            "此处添加文本",
            "此处添加副标题"
        };

        public static bool IsClearableType(int placeholderType)
        {
            return placeholderType == Title
                || placeholderType == Body
                || placeholderType == CenterTitle
                || placeholderType == Subtitle
                || placeholderType == VerticalTitle
                || placeholderType == VerticalBody;
        }

        public static string TypeName(int placeholderType)
        {
            switch (placeholderType)
            {
                case Title:
                case CenterTitle:
                case VerticalTitle:
                    return "title";
                case Subtitle:
                    return "subtitle";
                case Body:
                case VerticalBody:
                    return "body";
                default:
                    return "other";
            }
        }

        public static bool IsEmptyText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            string t = text.Replace("\r", "\n").Trim();
            foreach (string prompt in EmptyPrompts)
            {
                if (string.Equals(t, prompt, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void Remember(
            List<ClearedPlaceholderInfo> list,
            int placeholderType,
            int shapeComId)
        {
            if (list == null)
            {
                return;
            }

            list.Add(new ClearedPlaceholderInfo
            {
                PlaceholderType = TypeName(placeholderType),
                ShapeComId = shapeComId
            });
        }
    }
}
