using System;
using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    internal enum PptHtmlMissingShapeAction
    {
        Fail,
        Skip,
        Create
    }

    /// <summary>
    /// 导出 HTML 应用到空白/其它页：目标页找不到原 ShapeId 时的 upsert 规划。
    /// 旧 HTML 仍带 freeform/smartart/group/unknown → Skip（警告）；
    /// B2 起 read 已栅格为 picture，可走 Create。
    /// </summary>
    internal static class PptHtmlApplyUpsert
    {
        public const string DenyCreateMessage =
            "未允许新建：节点缺少本页 ShapeId（或目标页找不到）。要铺新形状请传 allow_create=true。";

        public static PptHtmlMissingShapeAction PlanMissingShape(
            PptHtmlApplyNode node,
            out string plannedType,
            out string error)
        {
            plannedType = null;
            error = null;
            if (node == null)
            {
                error = "无效节点";
                return PptHtmlMissingShapeAction.Fail;
            }

            string type = node.ShapeType ?? "";
            if (type.StartsWith("placeholder_", StringComparison.OrdinalIgnoreCase))
            {
                type = "textbox";
            }

            if (string.IsNullOrEmpty(type))
            {
                if (node.TableCells != null)
                {
                    type = "table";
                }
                else if (!string.IsNullOrWhiteSpace(node.DataSrc)
                    || !string.IsNullOrEmpty(node.ResolvedLocalPath))
                {
                    type = "picture";
                }
                else
                {
                    type = "textbox";
                }
            }

            if (type == "group")
            {
                if (node.Children == null || node.Children.Count == 0)
                {
                    plannedType = type;
                    return PptHtmlMissingShapeAction.Skip;
                }

                plannedType = type;
                return PptHtmlMissingShapeAction.Create;
            }

            if (IsSkipOnCreate(type) || !PptShapeTypeMap.IsCreatable(type))
            {
                plannedType = type;
                return PptHtmlMissingShapeAction.Skip;
            }

            if (!node.HasGeometry)
            {
                error = "目标页找不到 ShapeId=" + (node.ShapeId ?? "")
                    + "，新建还原须有 style 几何（left/top/width/height %）";
                return PptHtmlMissingShapeAction.Fail;
            }

            if ((type == "picture" || type == "media")
                && string.IsNullOrWhiteSpace(node.DataSrc)
                && string.IsNullOrEmpty(node.ResolvedLocalPath))
            {
                error = "目标页找不到 ShapeId=" + (node.ShapeId ?? "")
                    + "，新建 " + type + " 须有 data-src";
                return PptHtmlMissingShapeAction.Fail;
            }

            if (!TryValidateChartCreate(node, type, out error))
            {
                return PptHtmlMissingShapeAction.Fail;
            }

            plannedType = type;
            return PptHtmlMissingShapeAction.Create;
        }

        public static void MutateNodeForCreate(PptHtmlApplyNode node, string plannedType, List<string> warnings)
        {
            if (node == null)
            {
                return;
            }

            string originalId = node.ShapeId ?? "";
            string originalType = node.ShapeType ?? "";
            node.ShapeType = plannedType ?? "textbox";
            node.IsCreate = true;
            node.ShapeComId = null;
            node.ShapeId = null;
            warnings?.Add(
                "目标页无形状 " + originalId
                + (string.IsNullOrEmpty(originalType) ? "" : "（" + originalType + "）")
                + "，已按 " + node.ShapeType + " 新建还原");
            if (node.Children == null)
            {
                return;
            }

            foreach (PptHtmlApplyNode child in node.Children)
            {
                if (child == null)
                {
                    continue;
                }

                MutateNodeForCreate(child, child.ShapeType, warnings: null);
            }
        }

        public static void AddSkipWarning(PptHtmlApplyNode node, string plannedType, List<string> warnings)
        {
            warnings?.Add(
                "跳过无法新建的形状 " + (node?.ShapeId ?? "")
                + " type=" + (plannedType ?? node?.ShapeType ?? "")
                + "（如 freeform/smartart；其余形状继续写入）");
        }

        private static bool IsSkipOnCreate(string type)
        {
            return type == "freeform"
                || type == "smartart"
                || type == "unknown";
        }

        /// <summary>显式新建节点（无 ShapeId）预检。</summary>
        public static bool TryValidateExplicitCreate(PptHtmlApplyNode node, out string error)
        {
            error = null;
            if (node == null)
            {
                error = "无效节点";
                return false;
            }

            string type = node.ShapeType ?? "";
            if (type == "group")
            {
                if (node.Children == null || node.Children.Count == 0)
                {
                    error = "不能新建空 group";
                    return false;
                }

                if (!node.HasGeometry)
                {
                    error = "新建节点必须提供 style 几何";
                    return false;
                }

                return true;
            }

            if (IsSkipOnCreate(type) || !PptShapeTypeMap.IsCreatable(type))
            {
                // 显式新建不可建类型：跳过由调用方处理；这里标为可 Skip
                return true;
            }

            if (!node.HasGeometry)
            {
                error = "新建节点必须提供 style 几何";
                return false;
            }

            if ((type == "picture" || type == "media")
                && string.IsNullOrWhiteSpace(node.DataSrc)
                && string.IsNullOrEmpty(node.ResolvedLocalPath))
            {
                error = "新建 " + type + " 必须提供 data-src";
                return false;
            }

            return TryValidateChartCreate(node, type, out error);
        }

        private static bool TryValidateChartCreate(PptHtmlApplyNode node, string type, out string error)
        {
            error = null;
            if (type != "chart")
            {
                return true;
            }

            if (node.ChartGrid == null || !node.ChartGrid.IsPourable)
            {
                error = "新建 chart 必须在节点内嵌 <table> 灌数，不能建空图";
                return false;
            }

            return PptHtmlChartIo.TryParseType(node.ChartType, out _, out _, out error);
        }

        public static bool ShouldSkipExplicitCreate(PptHtmlApplyNode node)
        {
            string type = node?.ShapeType ?? "";
            if (type == "group")
            {
                return false;
            }

            return IsSkipOnCreate(type) || !PptShapeTypeMap.IsCreatable(type);
        }
    }
}
