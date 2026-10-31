using System;
using System.Collections.Generic;
using WordAddIn1.BrowserHost;

namespace BrowserNavigateTest
{
    internal static class ScriptDialogCases
    {
        public static int Run(string filter)
        {
            var cases = new List<Action>
            {
                Ac1,
                Ac2,
                Ac3,
                Ac4,
                Ac5,
                Ac6,
                Ac7,
                Ac8,
                Ac9,
                Ac10
            };

            int failed = 0;
            int ran = 0;
            foreach (Action act in cases)
            {
                string name = act.Method.Name;
                if (!string.Equals(filter, "script-dialog", StringComparison.OrdinalIgnoreCase)
                    && filter.IndexOf(name, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                ran++;
                try
                {
                    act();
                    Console.WriteLine("PASS " + name);
                }
                catch (Exception ex)
                {
                    failed++;
                    Console.WriteLine("FAIL " + name + " " + ex.Message);
                }
            }

            if (ran == 0)
            {
                Console.WriteLine("没有匹配的 script-dialog 用例: " + filter);
                return 1;
            }

            Console.WriteLine(failed == 0 ? "全部通过 " + ran : "失败 " + failed + "/" + ran);
            return failed == 0 ? 0 : 1;
        }

        private static void Ac1()
        {
            BrowserScriptDialogSession session = Open(BrowserScriptDialogKind.Confirm, "确定撤回该工时吗？", "");
            string tree = Tree(session);
            MustContain(tree, "- dialog confirm \"确定撤回该工时吗？\"");
            MustContain(tree, "button ref=e1 \"确定\"");
            MustContain(tree, "button ref=e2 \"取消\"");
            Equal(CountOf(tree, "button ref="), 2, "只有确定和取消");
            True(tree.IndexOf("row", StringComparison.Ordinal) < 0, "没有页面行");
            True(tree.IndexOf("textbox", StringComparison.Ordinal) < 0, "confirm 没有输入框");
        }

        private static void Ac2()
        {
            var sink = new Sink();
            BrowserScriptDialogSession session = Open(BrowserScriptDialogKind.Confirm, "确定撤回该工时吗？", "", sink);
            string channel = Install(session);
            BrowserScriptDialogStep step = BrowserScriptDialog.Apply(
                session, channel, "click", Ref(session, BrowserScriptDialog.ControlAccept), null, out string error);
            Equal(step, BrowserScriptDialogStep.Accept, error);
            BrowserScriptDialog.Commit(session, step);
            True(sink.Accepted, "已 Accept");
            True(sink.Completed, "deferral 已 Complete");
            True(!session.IsPending, "挂起被清掉");
            BrowserRefStore.Clear(channel);
        }

        private static void Ac3()
        {
            var sink = new Sink();
            BrowserScriptDialogSession session = Open(BrowserScriptDialogKind.Confirm, "确定撤回该工时吗？", "", sink);
            string channel = Install(session);
            BrowserScriptDialogStep step = BrowserScriptDialog.Apply(
                session, channel, "click", Ref(session, BrowserScriptDialog.ControlDismiss), null, out _);
            Equal(step, BrowserScriptDialogStep.Dismiss, "取消");
            BrowserScriptDialog.Commit(session, step);
            True(!sink.Accepted, "没有 Accept");
            True(sink.Completed, "deferral 已 Complete");
            True(!session.IsPending, "挂起被清掉");
            BrowserRefStore.Clear(channel);
        }

        private static void Ac4()
        {
            BrowserScriptDialogSession session = Open(BrowserScriptDialogKind.Confirm, "确定撤回该工时吗？", "");
            string channel = Install(session);
            BrowserScriptDialogStep step = BrowserScriptDialog.Apply(
                session, channel, "click", "e99", null, out string error);
            Equal(step, BrowserScriptDialogStep.Reject, "页面 ref");
            Equal(error, BrowserScriptDialog.MustHandleFirst, "文案");
            True(session.IsPending, "挂起仍在");
            True(BrowserRefStore.TryGet(channel, "e1", out _), "对话框 ref 还在");
            BrowserRefStore.Clear(channel);
        }

        private static void Ac5()
        {
            BrowserScriptDialogSession session = Open(BrowserScriptDialogKind.Confirm, "确定撤回该工时吗？", "");
            Equal(Tree(session), Tree(session), "再渲染仍是同一棵树");
        }

        private static void Ac6()
        {
            BrowserScriptDialogSession session = Open(BrowserScriptDialogKind.Confirm, "确定撤回该工时吗？", "");
            True(session.IsPending, "没人点就继续挂着");
            MustContain(Tree(session), "- dialog confirm \"确定撤回该工时吗？\"");
        }

        private static void Ac7()
        {
            var sink = new Sink();
            BrowserScriptDialogSession session = Open(BrowserScriptDialogKind.Alert, "已保存", "", sink);
            string tree = Tree(session);
            MustContain(tree, "- dialog alert \"已保存\"");
            MustContain(tree, "button ref=e1 \"确定\"");
            True(tree.IndexOf("取消", StringComparison.Ordinal) < 0, "alert 没有取消");
            string channel = Install(session);
            BrowserScriptDialogStep step = BrowserScriptDialog.Apply(
                session, channel, "click", "e1", null, out _);
            BrowserScriptDialog.Commit(session, step);
            True(!session.IsPending, "同意后挂起被清掉");
            True(sink.Accepted && sink.Completed, "alert 同意");
            BrowserRefStore.Clear(channel);
        }

        private static void Ac8()
        {
            var sink = new Sink();
            BrowserScriptDialogSession session = Open(BrowserScriptDialogKind.Prompt, "请输入原因", "甲", sink);
            MustContain(Tree(session), "value=\"甲\"");
            string channel = Install(session);
            string input = Ref(session, BrowserScriptDialog.ControlInput);
            BrowserScriptDialogStep typed = BrowserScriptDialog.Apply(
                session, channel, "type", input, "乙", out _);
            Equal(typed, BrowserScriptDialogStep.Typed, "type");
            True(!sink.Accepted, "type 不 Accept");
            MustContain(Tree(session), "value=\"乙\"");
            channel = Install(session);
            BrowserScriptDialogStep step = BrowserScriptDialog.Apply(
                session, channel, "click", Ref(session, BrowserScriptDialog.ControlAccept), null, out _);
            BrowserScriptDialog.Commit(session, step);
            Equal(sink.ResultText, "乙", "ResultText");
            True(sink.Accepted && sink.Completed, "再同意");
            BrowserRefStore.Clear(channel);
        }

        private static void Ac9()
        {
            BrowserScriptDialogSession session = Open(BrowserScriptDialogKind.Confirm, "你好", "");
            string tree = Tree(session);
            True(tree.StartsWith("- dialog confirm ", StringComparison.Ordinal), "只从对话框生成");
            True(tree.IndexOf("提交工时", StringComparison.Ordinal) < 0, "页内按钮不会被写成 dialog");
            True(tree.IndexOf("撤回", StringComparison.Ordinal) < 0, "没有页面上的撤回");
        }

        private static void Ac10()
        {
            BrowserScriptDialogSession session = Open(BrowserScriptDialogKind.Confirm, "确定删除吗？", "");
            BrowserScriptDialogView view = BrowserScriptDialog.Render(session);
            BrowserRefEntry accept = view.Refs["e1"];
            True(accept.IsScriptDialog, "对话框 ref");
            Equal(accept.Name, "确定", "名字是确定");
            Equal(accept.ScriptDialogControl, BrowserScriptDialog.ControlAccept, "accept");
            True(!accept.BackendDomNodeId.HasValue, "没有 DOM 节点");
            True(accept.Name.IndexOf("删除", StringComparison.Ordinal) < 0, "不带删除标记");
            string channel = "browser:agent:script-dialog-ac10";
            BrowserRefStore.Replace(channel, view.Refs);
            BrowserScriptDialogStep step = BrowserScriptDialog.Apply(
                session, channel, "click", "e1", null, out _);
            Equal(step, BrowserScriptDialogStep.Accept, "同意在风险词之前返回");
            BrowserRefStore.Clear(channel);
        }

        private sealed class Sink
        {
            public bool Accepted { get; private set; }

            public bool Completed { get; private set; }

            public string ResultText { get; private set; }

            public void Accept(string text)
            {
                Accepted = true;
                ResultText = text ?? "";
            }

            public void Complete()
            {
                Completed = true;
            }
        }

        private static BrowserScriptDialogSession Open(
            BrowserScriptDialogKind kind,
            string message,
            string initial)
        {
            return Open(kind, message, initial, new Sink());
        }

        private static BrowserScriptDialogSession Open(
            BrowserScriptDialogKind kind,
            string message,
            string initial,
            Sink sink)
        {
            return BrowserScriptDialogSession.Create(
                kind,
                message,
                initial,
                sink.Accept,
                sink.Complete);
        }

        private static string Tree(BrowserScriptDialogSession session)
        {
            return BrowserScriptDialog.Render(session).TreeText.Replace("\r\n", "\n");
        }

        private static string Install(BrowserScriptDialogSession session)
        {
            string channel = "browser:agent:script-dialog-" + Guid.NewGuid().ToString("N");
            BrowserRefStore.Replace(channel, BrowserScriptDialog.Render(session).Refs);
            return channel;
        }

        private static string Ref(BrowserScriptDialogSession session, string control)
        {
            foreach (KeyValuePair<string, BrowserRefEntry> pair in BrowserScriptDialog.Render(session).Refs)
            {
                if (pair.Value.ScriptDialogControl == control)
                {
                    return pair.Key;
                }
            }

            throw new InvalidOperationException("没有 " + control);
        }

        private static void MustContain(string tree, string needle)
        {
            if (tree.IndexOf(needle, StringComparison.Ordinal) < 0)
            {
                throw new InvalidOperationException("缺少 " + needle + "\n" + tree);
            }
        }

        private static void Equal(object actual, object expected, string label)
        {
            if (!object.Equals(actual, expected))
            {
                throw new InvalidOperationException(label + " 实际=" + actual + " 期望=" + expected);
            }
        }

        private static void True(bool condition, string label)
        {
            if (!condition)
            {
                throw new InvalidOperationException(label);
            }
        }

        private static int CountOf(string tree, string needle)
        {
            int count = 0;
            int at = 0;
            while (true)
            {
                int found = tree.IndexOf(needle, at, StringComparison.Ordinal);
                if (found < 0)
                {
                    return count;
                }

                count++;
                at = found + needle.Length;
            }
        }
    }
}
