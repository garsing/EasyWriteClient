using System;
using System.Collections.Generic;
using System.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>D18 启发式：密码 / 提交登录支付 / 删除。</summary>
    internal static class BrowserRiskGuard
    {
        private static readonly string[] LoginSubmitPayMarkers =
        {
            "登录", "登陆", "立即登录", "注册并登录",
            "sign in", "log in", "login", "signin",
            "提交", "确认提交", "submit",
            "支付", "付款", "结算", "下单", "确认支付", "验证并支付",
            "checkout", "pay now", "place order", "payment"
        };

        private static readonly string[] DeleteMarkers =
        {
            "删除", "delete", "注销", "清空账号", "delete account", "remove account"
        };

        private static readonly string[] PasswordNameMarkers =
        {
            "密码", "口令", "password", "passwd", "pwd"
        };

        public static string CheckType(BrowserRefEntry entry, DomNodeProbe probe)
        {
            if (IsPasswordControl(entry, probe))
            {
                return "已拒绝向密码框输入；请用户在看得见的窗里自己填写密码";
            }

            return null;
        }

        public static string CheckClick(BrowserRefEntry entry, DomNodeProbe probe)
        {
            if (IsDeleteControl(entry, probe))
            {
                return "已拒绝点击删除类控件；请用户自己操作";
            }

            if (IsLoginSubmitPayControl(entry, probe))
            {
                return "已拒绝代点登录/提交/支付；请用户在看得见的窗里自己点击";
            }

            return null;
        }

        public static string CheckPressEnter(BrowserRefEntry focusEntry, DomNodeProbe probe, bool pageHasPasswordHint)
        {
            if (IsPasswordControl(focusEntry, probe))
            {
                return "已拒绝在密码框按 Enter（可能代提交）；请用户自己操作";
            }

            if (IsLoginSubmitPayControl(focusEntry, probe))
            {
                return "已拒绝用 Enter 代提交/登录/支付；请用户自己操作";
            }

            // 焦点在普通输入框，但同页像登录表（有 password）→ 拒绝 Enter 代提交
            if (pageHasPasswordHint && LooksLikeTextInput(focusEntry, probe))
            {
                return "已拒绝在含密码表单中按 Enter 代提交；请用户自己操作";
            }

            return null;
        }

        public static bool IsPasswordControl(BrowserRefEntry entry, DomNodeProbe probe)
        {
            if (probe != null)
            {
                if (string.Equals(probe.InputType, "password", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (!string.IsNullOrEmpty(probe.Autocomplete)
                    && probe.Autocomplete.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            string name = CombinedLabel(entry, probe);
            if (LooksLikeTextInput(entry, probe) && ContainsAny(name, PasswordNameMarkers))
            {
                return true;
            }

            return false;
        }

        public static bool IsLoginSubmitPayControl(BrowserRefEntry entry, DomNodeProbe probe)
        {
            string label = CombinedLabel(entry, probe);
            if (ContainsAny(label, LoginSubmitPayMarkers))
            {
                return true;
            }

            if (probe != null
                && string.Equals(probe.InputType, "submit", StringComparison.OrdinalIgnoreCase)
                && string.IsNullOrWhiteSpace(label))
            {
                // type=submit 且无名：交给调用方结合 pageHasPassword 再判；此处先不单独拒绝
                return false;
            }

            return false;
        }

        public static bool IsDeleteControl(BrowserRefEntry entry, DomNodeProbe probe)
        {
            string role = entry?.Role ?? probe?.Role ?? "";
            bool buttonLike = string.Equals(role, "button", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "link", StringComparison.OrdinalIgnoreCase)
                || string.Equals(probe?.Tag, "BUTTON", StringComparison.OrdinalIgnoreCase)
                || string.Equals(probe?.Tag, "A", StringComparison.OrdinalIgnoreCase);

            if (!buttonLike && probe == null && entry == null)
            {
                return false;
            }

            // 方案 A：凡带删除/Delete 的 button/link 一律拒
            if (!buttonLike)
            {
                // 无名 role 但文本像删除按钮
                buttonLike = ContainsAny(CombinedLabel(entry, probe), DeleteMarkers);
            }

            if (!buttonLike)
            {
                return false;
            }

            return ContainsAny(CombinedLabel(entry, probe), DeleteMarkers);
        }

        public static bool IsFrameRole(BrowserRefEntry entry)
        {
            if (entry == null || string.IsNullOrEmpty(entry.Role))
            {
                return false;
            }

            return string.Equals(entry.Role, "Iframe", StringComparison.OrdinalIgnoreCase)
                || string.Equals(entry.Role, "iframe", StringComparison.OrdinalIgnoreCase)
                || string.Equals(entry.Role, "Frame", StringComparison.OrdinalIgnoreCase)
                || string.Equals(entry.Role, "frame", StringComparison.OrdinalIgnoreCase);
        }

        private static bool LooksLikeTextInput(BrowserRefEntry entry, DomNodeProbe probe)
        {
            string role = entry?.Role ?? "";
            if (string.Equals(role, "textbox", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "searchbox", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (probe != null && string.Equals(probe.Tag, "INPUT", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private static string CombinedLabel(BrowserRefEntry entry, DomNodeProbe probe)
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(entry?.Name))
            {
                parts.Add(entry.Name);
            }

            if (probe != null)
            {
                if (!string.IsNullOrWhiteSpace(probe.AccessibleName))
                {
                    parts.Add(probe.AccessibleName);
                }

                if (!string.IsNullOrWhiteSpace(probe.ValueAttr))
                {
                    parts.Add(probe.ValueAttr);
                }

                if (!string.IsNullOrWhiteSpace(probe.InnerText))
                {
                    parts.Add(probe.InnerText);
                }
            }

            return string.Join(" ", parts);
        }

        private static bool ContainsAny(string haystack, IEnumerable<string> needles)
        {
            if (string.IsNullOrWhiteSpace(haystack))
            {
                return false;
            }

            string h = haystack.Trim().ToLowerInvariant();
            return needles.Any(n => !string.IsNullOrEmpty(n) && h.IndexOf(n.ToLowerInvariant(), StringComparison.Ordinal) >= 0);
        }
    }

    internal sealed class DomNodeProbe
    {
        public string Tag { get; set; }
        public string InputType { get; set; }
        public string Autocomplete { get; set; }
        public string Role { get; set; }
        public string AccessibleName { get; set; }
        public string ValueAttr { get; set; }
        public string InnerText { get; set; }
        public bool PageHasPasswordInput { get; set; }
    }
}
