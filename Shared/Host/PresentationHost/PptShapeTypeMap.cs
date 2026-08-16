using System;
using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    /// <summary>
    /// COM Shape.Type / AutoShapeType / PlaceholderFormat.Type → data-shape-type（§5.3 加长名单）。
    /// </summary>
    internal static class PptShapeTypeMap
    {
        // MsoShapeType
        private const int MsoAutoShape = 1;
        private const int MsoCallout = 2;
        private const int MsoChart = 3;
        private const int MsoFreeform = 5;
        private const int MsoGroup = 6;
        private const int MsoLine = 9;
        private const int MsoLinkedPicture = 11;
        private const int MsoPicture = 13;
        private const int MsoPlaceholder = 14;
        private const int MsoMedia = 16;
        private const int MsoTextBox = 17;
        private const int MsoTable = 19;
        private const int MsoDiagram = 21;
        private const int MsoSmartArt = 24;

        // PpPlaceholderType
        private const int PpTitle = 1;
        private const int PpBody = 2;
        private const int PpCenterTitle = 3;
        private const int PpSubtitle = 4;
        private const int PpVerticalTitle = 16;
        private const int PpVerticalBody = 17;

        private static readonly Dictionary<int, string> AutoShapeNames = BuildAutoShapeNames();

        public static string FromShapeType(
            int shapeType,
            int? autoShapeType,
            int? placeholderType)
        {
            switch (shapeType)
            {
                case MsoPlaceholder:
                    return FromPlaceholder(placeholderType);
                case MsoTextBox:
                    return "textbox";
                case MsoPicture:
                case MsoLinkedPicture:
                    return "picture";
                case MsoTable:
                    return "table";
                case MsoChart:
                    return "chart";
                case MsoSmartArt:
                case MsoDiagram:
                    return "smartart";
                case MsoGroup:
                    return "group";
                case MsoMedia:
                    return "media";
                case MsoFreeform:
                    return "freeform";
                case MsoLine:
                    return "line";
                case MsoCallout:
                    if (autoShapeType.HasValue
                        && AutoShapeNames.TryGetValue(autoShapeType.Value, out string calloutName))
                    {
                        return calloutName;
                    }

                    return "callout_1";
                case MsoAutoShape:
                    if (autoShapeType.HasValue
                        && AutoShapeNames.TryGetValue(autoShapeType.Value, out string name))
                    {
                        return name;
                    }

                    return "unknown";
                default:
                    if (autoShapeType.HasValue
                        && AutoShapeNames.TryGetValue(autoShapeType.Value, out string fallback))
                    {
                        return fallback;
                    }

                    return "unknown";
            }
        }

        public static string PreferTag(string shapeType, bool hasText)
        {
            if (string.IsNullOrEmpty(shapeType))
            {
                return hasText ? "div" : "div";
            }

            if (shapeType == "placeholder_title")
            {
                return "h1";
            }

            if (shapeType == "picture")
            {
                return "img";
            }

            if (shapeType == "table")
            {
                return "table";
            }

            return "div";
        }

        public static bool IsNonEditable(string shapeType)
        {
            return shapeType == "picture"
                || shapeType == "chart"
                || shapeType == "smartart"
                || shapeType == "media";
        }

        public static bool TryGetAutoShapeType(string shapeType, out int autoShapeType)
        {
            autoShapeType = 0;
            if (string.IsNullOrEmpty(shapeType))
            {
                return false;
            }

            EnsureReverseMap();
            return NameToAutoShape.TryGetValue(shapeType, out autoShapeType);
        }

        public static bool IsCreatable(string shapeType)
        {
            if (string.IsNullOrEmpty(shapeType))
            {
                return false;
            }

            if (shapeType.StartsWith("placeholder_", StringComparison.Ordinal)
                || shapeType == "smartart"
                || shapeType == "group"
                || shapeType == "unknown"
                || shapeType == "freeform")
            {
                return false;
            }

            if (shapeType == "textbox"
                || shapeType == "table"
                || shapeType == "picture"
                || shapeType == "chart"
                || shapeType == "media"
                || shapeType == "line")
            {
                return true;
            }

            return TryGetAutoShapeType(shapeType, out _);
        }

        private static Dictionary<string, int> NameToAutoShape;

        private static void EnsureReverseMap()
        {
            if (NameToAutoShape != null)
            {
                return;
            }

            NameToAutoShape = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (KeyValuePair<int, string> kv in AutoShapeNames)
            {
                if (!NameToAutoShape.ContainsKey(kv.Value))
                {
                    NameToAutoShape[kv.Value] = kv.Key;
                }
            }
        }

        private static string FromPlaceholder(int? placeholderType)
        {
            if (!placeholderType.HasValue)
            {
                return "placeholder_other";
            }

            int t = placeholderType.Value;
            if (t == PpTitle || t == PpCenterTitle || t == PpVerticalTitle)
            {
                return "placeholder_title";
            }

            if (t == PpBody || t == PpSubtitle || t == PpVerticalBody)
            {
                return "placeholder_body";
            }

            return "placeholder_other";
        }

        private static Dictionary<int, string> BuildAutoShapeNames()
        {
            // MsoAutoShapeType 常见值 → §5.3
            var map = new Dictionary<int, string>
            {
                [1] = "rectangle",
                [2] = "parallelogram",
                [3] = "trapezoid",
                [4] = "diamond",
                [5] = "rounded_rectangle",
                [6] = "octagon",
                [7] = "triangle",
                [8] = "right_triangle",
                [9] = "oval",
                [10] = "hexagon",
                [11] = "cross",
                [12] = "pentagon",
                [13] = "can",
                [14] = "cube",
                [15] = "bevel",
                [18] = "donut",
                [19] = "no_symbol",
                [20] = "block_arc",
                [21] = "heart",
                [22] = "lightning_bolt",
                [23] = "sun",
                [24] = "moon",
                [25] = "arc",
                [26] = "double_bracket",
                [27] = "double_brace",
                [28] = "plaque",
                [29] = "left_bracket",
                [30] = "right_bracket",
                [31] = "left_brace",
                [32] = "right_brace",
                [33] = "right_arrow",
                [34] = "left_arrow",
                [35] = "up_arrow",
                [36] = "down_arrow",
                [37] = "left_right_arrow",
                [38] = "up_down_arrow",
                [39] = "quad_arrow",
                [40] = "left_right_up_arrow",
                [41] = "bent_arrow",
                [42] = "uturn_arrow",
                [43] = "left_up_arrow",
                [44] = "bent_up_arrow",
                [45] = "curved_right_arrow",
                [46] = "curved_left_arrow",
                [47] = "curved_up_arrow",
                [48] = "curved_down_arrow",
                [49] = "striped_right_arrow",
                [50] = "notched_right_arrow",
                [51] = "home_plate",
                [52] = "chevron",
                [53] = "right_arrow_callout",
                [54] = "left_arrow_callout",
                [55] = "up_arrow_callout",
                [56] = "down_arrow_callout",
                [57] = "left_right_arrow_callout",
                [58] = "up_down_arrow_callout",
                [59] = "quad_arrow_callout",
                [61] = "flowchart_process",
                [62] = "flowchart_alternate_process",
                [63] = "flowchart_decision",
                [64] = "flowchart_data",
                [65] = "flowchart_predefined_process",
                [66] = "flowchart_internal_storage",
                [67] = "flowchart_document",
                [68] = "flowchart_multidocument",
                [69] = "flowchart_terminator",
                [70] = "flowchart_preparation",
                [71] = "flowchart_manual_input",
                [72] = "flowchart_manual_operation",
                [73] = "flowchart_connector",
                [74] = "flowchart_offpage_connector",
                [75] = "flowchart_card",
                [76] = "flowchart_punched_tape",
                [77] = "flowchart_summing_junction",
                [78] = "flowchart_or",
                [79] = "flowchart_collate",
                [80] = "flowchart_sort",
                [81] = "flowchart_extract",
                [82] = "flowchart_merge",
                [83] = "flowchart_stored_data",
                [84] = "flowchart_delay",
                [85] = "flowchart_sequential_access_storage",
                [86] = "flowchart_magnetic_disk",
                [87] = "flowchart_direct_access_storage",
                [88] = "flowchart_display",
                [89] = "explosion_1",
                [90] = "explosion_2",
                [91] = "star_4",
                [92] = "star_5",
                [93] = "star_8",
                [94] = "star_16",
                [95] = "star_24",
                [96] = "star_32",
                [97] = "irregular_seal_1",
                [98] = "irregular_seal_2",
                [99] = "ribbon",
                [100] = "ribbon_2",
                [101] = "ellipse_ribbon",
                [102] = "ellipse_ribbon_2",
                [103] = "vertical_scroll",
                [104] = "horizontal_scroll",
                [105] = "wave",
                [106] = "double_wave",
                [107] = "rectangular_callout", // map near wedge
                [108] = "rounded_rectangular_callout",
                [109] = "oval_callout",
                [110] = "cloud_callout",
                [111] = "wedge_rect_callout",
                [112] = "wedge_round_rect_callout",
                [113] = "wedge_ellipse_callout",
                [118] = "border_callout_1",
                [119] = "border_callout_2",
                [120] = "border_callout_3",
                [121] = "accent_callout_1",
                [122] = "accent_callout_2",
                [123] = "accent_callout_3",
                [124] = "callout_1",
                [125] = "callout_2",
                [126] = "callout_3",
                [127] = "accent_border_callout_1",
                [128] = "accent_border_callout_2",
                [129] = "accent_border_callout_3",
                [131] = "heptagon",
                [132] = "decagon",
                [133] = "dodecagon",
                [134] = "star_6",
                [135] = "star_7",
                [136] = "star_10",
                [137] = "star_12",
                [138] = "round_1_rectangle",
                [139] = "round_2_same_rectangle",
                [140] = "round_2_diag_rectangle",
                [141] = "snip_round_rectangle",
                [142] = "snip_1_rectangle",
                [143] = "snip_2_same_rectangle",
                [144] = "snip_2_diag_rectangle",
                [145] = "tear",
                [158] = "chord",
                [183] = "pie",
                [184] = "cloud"
            };
            return map;
        }
    }
}
