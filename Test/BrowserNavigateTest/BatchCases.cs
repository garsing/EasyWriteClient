using System;
using System.Collections;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using WordAddIn1;
using WordAddIn1.BrowserHost;

namespace BrowserNavigateTest
{
    /// <summary>F_browser_interact actions[]。前缀 bat- / bat-unit-。</summary>
    internal static class BatchCases
    {
        public static int RunUnit()
        {
            Console.WriteLine("浏览器 一次操作多个节点（bat-unit，不连 Desktop）");
            Console.WriteLine();
            var run = new TestRun();
            CompileEmpty(run);
            CompileMissing(run);
            CompileMix(run);
            CompileKeepsBadItem(run);
            MsgAllOk(run);
            MsgPartial(run);
            MsgAllFail(run);
            UrlHash(run);
            CompileNotArray(run);
            CompileItemNotObject(run);
            CompileSingleTop(run);
            Console.WriteLine();
            Console.WriteLine("正式用例 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped
                + "  （目标 11）");
            return run.CasesFailed == 0 ? 0 : 1;
        }

        public static void Run(TestRun run, NavSession s, FixtureHttp http)
        {
            RunContract(run, s, http);
            RunHappy(run, s, http);
            RunPartial(run, s, http);
            RunStop(run, s, http);
        }

        private static void CompileEmpty(TestRun run)
        {
            List<BrowserInteractStep> steps;
            string err;
            bool ok = BrowserInteractBatch.TryCompile(
                new Dictionary<string, object> { ["actions"] = new List<object>() },
                out steps,
                out err);
            IntCases.Ok(
                run,
                "bat-unit-empty-actions",
                !ok && (err ?? "").IndexOf("不能为空", StringComparison.Ordinal) >= 0,
                err ?? "应契约失败");
        }

        private static void CompileMissing(TestRun run)
        {
            List<BrowserInteractStep> steps;
            string err;
            bool ok = BrowserInteractBatch.TryCompile(new Dictionary<string, object>(), out steps, out err);
            IntCases.Ok(
                run,
                "bat-unit-no-action",
                !ok && (err ?? "").IndexOf("action", StringComparison.OrdinalIgnoreCase) >= 0,
                err ?? "应契约失败");
        }

        private static void CompileMix(TestRun run)
        {
            var args = new Dictionary<string, object>
            {
                ["action"] = "click",
                ["ref"] = "e9",
                ["actions"] = new List<object>
                {
                    new Dictionary<string, object>
                    {
                        ["action"] = "type",
                        ["ref"] = "e1",
                        ["text"] = "甲",
                    },
                    new Dictionary<string, object>
                    {
                        ["action"] = "type",
                        ["ref"] = "e2",
                        ["text"] = "乙",
                    },
                },
            };
            List<BrowserInteractStep> steps;
            string err;
            bool ok = BrowserInteractBatch.TryCompile(args, out steps, out err);
            IntCases.Ok(
                run,
                "bat-unit-mix-uses-array",
                ok
                    && steps != null
                    && steps.Count == 2
                    && string.Equals(steps[0].Action, "type", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(steps[0].Ref, "e1", StringComparison.Ordinal)
                    && string.Equals(steps[1].Text, "乙", StringComparison.Ordinal),
                err ?? (ok ? "混入参未只用数组" : "编译失败"));
        }

        private static void CompileKeepsBadItem(TestRun run)
        {
            var args = new Dictionary<string, object>
            {
                ["actions"] = new List<object>
                {
                    new Dictionary<string, object>
                    {
                        ["action"] = "type",
                        ["ref"] = "e1",
                        ["text"] = "甲",
                    },
                    new Dictionary<string, object>
                    {
                        ["action"] = "type",
                        ["ref"] = "e2",
                    },
                },
            };
            List<BrowserInteractStep> steps;
            string err;
            bool ok = BrowserInteractBatch.TryCompile(args, out steps, out err);
            IntCases.Ok(
                run,
                "bat-unit-compile-keeps-bad-item",
                ok && steps != null && steps.Count == 2 && string.IsNullOrEmpty(steps[1].Text),
                "缺 text 应仍编进列表，执行时该项失败");
        }

        private static void MsgAllOk(TestRun run)
        {
            var items = new List<BrowserInteractItemResult>
            {
                Item(0, "type", "success", null),
                Item(1, "click", "success", null),
            };
            string msg = BrowserInteractBatch.BuildMessage(items);
            IntCases.Ok(
                run,
                "bat-unit-msg-all-ok",
                msg != null && msg.StartsWith("全部成功", StringComparison.Ordinal),
                msg);
        }

        private static void MsgPartial(TestRun run)
        {
            var items = new List<BrowserInteractItemResult>
            {
                Item(0, "type", "success", null),
                Item(1, "type", "failed", "过期"),
                Item(2, "type", "skipped", "前一项改变了页面"),
            };
            string msg = BrowserInteractBatch.BuildMessage(items);
            IntCases.Ok(
                run,
                "bat-unit-msg-partial",
                msg != null
                    && msg.StartsWith("部分成功", StringComparison.Ordinal)
                    && msg.IndexOf("action[0]", StringComparison.Ordinal) >= 0
                    && msg.IndexOf("未执行", StringComparison.Ordinal) >= 0,
                msg);
        }

        private static void MsgAllFail(TestRun run)
        {
            var items = new List<BrowserInteractItemResult>
            {
                Item(0, "click", "failed", "过期"),
                Item(1, "click", "failed", "过期"),
            };
            string msg = BrowserInteractBatch.BuildMessage(items);
            IntCases.Ok(
                run,
                "bat-unit-msg-all-fail",
                msg != null && msg.StartsWith("全部失败", StringComparison.Ordinal),
                msg);
        }

        private static void UrlHash(TestRun run)
        {
            IntCases.Ok(
                run,
                "bat-unit-url-hash",
                !BrowserInteractBatch.UrlChanged("http://127.0.0.1/a#x", "http://127.0.0.1/a#y")
                    && BrowserInteractBatch.UrlChanged("http://127.0.0.1/a", "http://127.0.0.1/b"),
                "hash 变化不应算跳转");
        }

        private static void CompileNotArray(TestRun run)
        {
            List<BrowserInteractStep> steps;
            string err;
            bool ok = BrowserInteractBatch.TryCompile(
                new Dictionary<string, object> { ["actions"] = "not-array" },
                out steps,
                out err);
            IntCases.Ok(
                run,
                "bat-unit-actions-not-array",
                !ok && (err ?? "").IndexOf("对象数组", StringComparison.Ordinal) >= 0,
                err ?? "应契约失败");
        }

        private static void CompileItemNotObject(TestRun run)
        {
            List<BrowserInteractStep> steps;
            string err;
            bool ok = BrowserInteractBatch.TryCompile(
                new Dictionary<string, object>
                {
                    ["actions"] = new List<object> { "click" },
                },
                out steps,
                out err);
            IntCases.Ok(
                run,
                "bat-unit-item-not-object",
                !ok && (err ?? "").IndexOf("每一项须为对象", StringComparison.Ordinal) >= 0,
                err ?? "应契约失败");
        }

        private static void CompileSingleTop(TestRun run)
        {
            List<BrowserInteractStep> steps;
            string err;
            bool ok = BrowserInteractBatch.TryCompile(
                new Dictionary<string, object>
                {
                    ["action"] = "click",
                    ["ref"] = "e4",
                },
                out steps,
                out err);
            IntCases.Ok(
                run,
                "bat-unit-single-top",
                ok
                    && steps != null
                    && steps.Count == 1
                    && string.Equals(steps[0].Action, "click", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(steps[0].Ref, "e4", StringComparison.Ordinal),
                err ?? "单项顶层应编成 1 步");
        }

        private static BrowserInteractItemResult Item(int index, string action, string status, string error)
        {
            return new BrowserInteractItemResult
            {
                ActionIndex = index,
                Action = action,
                Status = status,
                Error = error,
            };
        }

        private static void RunContract(TestRun run, NavSession s, FixtureHttp http)
        {
            ToolResult noCh = s.InteractBatch(
                null,
                new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object> { ["action"] = "click", ["ref"] = "e1" },
                });
            if (noCh != null && !noCh.Success)
            {
                string err = noCh.Error ?? "";
                IntCases.Ok(
                    run,
                    "bat-neg-no-channel",
                    IntCases.ContainsAny(err, "渠道", "打开网页", "ref"),
                    err);
            }
            else
            {
                IntCases.Ok(run, "bat-neg-no-channel", noCh != null && noCh.Success, "Desktop 默认已是浏览器渠");
            }

            IntCases.ExpectFail(
                run,
                "bat-neg-wrong-kind",
                s.InteractBatch("w1", new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object> { ["action"] = "click", ["ref"] = "e1" },
                }),
                "渠道");
            IntCases.ExpectFail(
                run,
                "bat-neg-unknown-id",
                s.InteractBatch("b999", new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object> { ["action"] = "click", ["ref"] = "e1" },
                }),
                "渠道");

            string id = IntCases.OpenOk(run, s, http.Url("batch-form", "index.html"), "bat-neg-empty-actions");
            if (id == null)
            {
                IntCases.FailMany(
                    run,
                    "依赖 batch-form",
                    "bat-neg-empty-actions",
                    "bat-neg-no-action-or-actions",
                    "bat-neg-actions-string",
                    "bat-mix-top-ignored");
                return;
            }

            IntCases.ExpectFail(
                run,
                "bat-neg-empty-actions",
                s.InteractBatch(id, new List<Dictionary<string, object>>()),
                "不能为空");
            IntCases.ExpectFail(
                run,
                "bat-neg-no-action-or-actions",
                s.Interact(id),
                "action");
            IntCases.ExpectFail(
                run,
                "bat-neg-actions-string",
                s.InteractRaw(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["actions"] = "not-array",
                }),
                "对象数组");

            string tree = IntCases.WaitTree(s, id, true, "姓名", "提交");
            string name = IntCases.FindRef(tree, "姓名");
            string org = IntCases.FindRef(tree, "单位");
            string email = IntCases.FindRef(tree, "邮箱");
            string submit = IntCases.FindRef(tree, "提交", "button");
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(org) || string.IsNullOrEmpty(email) || string.IsNullOrEmpty(submit))
            {
                run.CaseFail("bat-mix-top-ignored", "缺姓名/单位/邮箱/提交 ref");
                return;
            }

            ToolResult mixed = s.InteractBatch(
                id,
                new List<Dictionary<string, object>>
                {
                    TypeRow(name, "甲"),
                    TypeRow(org, "乙所"),
                    TypeRow(email, "c@d.com"),
                },
                topAction: "click",
                topRef: submit);
            string snap = IntCases.Tree(mixed);
            bool pass = mixed != null
                && mixed.Success
                && Outcome(mixed) == "all_succeeded"
                && Rows(mixed).Count == 3
                && snap.IndexOf("已提交", StringComparison.Ordinal) < 0
                && ContainsAll(snap, "甲", "乙所", "c@d.com");
            IntCases.Ok(run, "bat-mix-top-ignored", pass, Dump(mixed, "顶层 click 不应执行"));
        }

        private static void RunHappy(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = IntCases.OpenOk(run, s, http.Url("batch-form", "index.html"), "bat-type-three");
            if (id == null)
            {
                IntCases.FailMany(
                    run,
                    "依赖 batch-form",
                    "bat-type-three",
                    "bat-results-order",
                    "bat-no-leak",
                    "bat-single-compat",
                    "bat-mix-verbs",
                    "bat-type-twice-overwrite",
                    "bat-click-checkbox",
                    "bat-hash-not-stop");
                return;
            }

            string tree = IntCases.WaitTree(s, id, true, "姓名", "课题类型");
            string name = IntCases.FindRef(tree, "姓名");
            string org = IntCases.FindRef(tree, "单位");
            string email = IntCases.FindRef(tree, "邮箱");
            string kind = IntCases.FindRef(tree, "课题类型");
            string submit = IntCases.FindRef(tree, "提交", "button");
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(org) || string.IsNullOrEmpty(email))
            {
                IntCases.FailMany(
                    run,
                    "缺输入框 ref",
                    "bat-type-three",
                    "bat-results-order",
                    "bat-no-leak",
                    "bat-single-compat",
                    "bat-mix-verbs",
                    "bat-type-twice-overwrite",
                    "bat-click-checkbox",
                    "bat-hash-not-stop");
                return;
            }

            ToolResult three = s.InteractBatch(
                id,
                new List<Dictionary<string, object>>
                {
                    TypeRow(name, "张三"),
                    TypeRow(org, "某研究所"),
                    TypeRow(email, "a@b.com"),
                });
            bool threeOk = three != null
                && three.Success
                && Outcome(three) == "all_succeeded"
                && FirstLine(three).StartsWith("全部成功", StringComparison.Ordinal)
                && AllStatus(three, "success")
                && Rows(three).Count == 3
                && ContainsAll(IntCases.Tree(three), "张三", "某研究所", "a@b.com");
            IntCases.Ok(run, "bat-type-three", threeOk, Dump(three, "三框未一次填入"));
            IntCases.Ok(
                run,
                "bat-results-order",
                threeOk
                    && IntAt(three, 0, "action_index") == 0
                    && IntAt(three, 1, "action_index") == 1
                    && IntAt(three, 2, "action_index") == 2
                    && string.Equals(Cell(Rows(three)[0], "action"), "type", StringComparison.OrdinalIgnoreCase),
                Dump(three, "results 顺序或 action_index 不对"));
            IntCases.Ok(
                run,
                "bat-no-leak",
                threeOk && !IntCases.HasLeak(three),
                "回包泄漏 backendNodeId / cssPath / HTML");

            ToolResult single = s.Interact(id, "type", name, text: "李四");
            bool singleOk = single != null
                && single.Success
                && Outcome(single) == "all_succeeded"
                && Rows(single).Count == 1
                && StatusAt(single, 0) == "success"
                && IntCases.Tree(single).IndexOf("李四", StringComparison.Ordinal) >= 0;
            IntCases.Ok(run, "bat-single-compat", singleOk, Dump(single, "单项未走 batch_outcome"));

            if (string.IsNullOrEmpty(kind) || string.IsNullOrEmpty(submit))
            {
                IntCases.FailMany(
                    run,
                    "缺课题类型或提交 ref",
                    "bat-mix-verbs",
                    "bat-type-twice-overwrite",
                    "bat-click-checkbox",
                    "bat-hash-not-stop");
                return;
            }

            tree = IntCases.WaitTree(s, id, true, "姓名", "提交");
            name = IntCases.FindRef(tree, "姓名") ?? name;
            kind = IntCases.FindRef(tree, "课题类型") ?? kind;
            submit = IntCases.FindRef(tree, "提交", "button") ?? submit;
            ToolResult mix = s.InteractBatch(
                id,
                new List<Dictionary<string, object>>
                {
                    TypeRow(name, "王五"),
                    new Dictionary<string, object>
                    {
                        ["action"] = "select",
                        ["ref"] = kind,
                        ["option"] = "应用研究",
                    },
                    new Dictionary<string, object>
                    {
                        ["action"] = "click",
                        ["ref"] = submit,
                    },
                });
            string mixTree = IntCases.Tree(mix);
            bool mixOk = mix != null
                && mix.Success
                && Outcome(mix) == "all_succeeded"
                && Rows(mix).Count == 3
                && mixTree.IndexOf("王五", StringComparison.Ordinal) >= 0
                && mixTree.IndexOf("已提交", StringComparison.Ordinal) >= 0;
            IntCases.Ok(run, "bat-mix-verbs", mixOk, Dump(mix, "混动词未全部成功"));

            tree = IntCases.WaitTree(s, id, true, "姓名", "同意条款", "页内锚点");
            name = IntCases.FindRef(tree, "姓名") ?? name;
            string agree = IntCases.FindRef(tree, "同意条款");
            string anchor = IntCases.FindRef(tree, "页内锚点");
            ToolResult twice = s.InteractBatch(
                id,
                new List<Dictionary<string, object>>
                {
                    TypeRow(name, "初值"),
                    TypeRow(name, "终值"),
                });
            string twiceTree = IntCases.Tree(twice);
            IntCases.Ok(
                run,
                "bat-type-twice-overwrite",
                twice != null
                    && twice.Success
                    && twiceTree.IndexOf("终值", StringComparison.Ordinal) >= 0,
                Dump(twice, "同一 ref 第二次 type 应覆盖"));

            if (string.IsNullOrEmpty(agree))
            {
                run.CaseFail("bat-click-checkbox", "缺同意条款 ref");
            }
            else
            {
                ToolResult ck = s.InteractBatch(
                    id,
                    new List<Dictionary<string, object>>
                    {
                        TypeRow(name, "勾选人"),
                        new Dictionary<string, object> { ["action"] = "click", ["ref"] = agree },
                    });
                IntCases.Ok(
                    run,
                    "bat-click-checkbox",
                    ck != null
                        && ck.Success
                        && IntCases.Tree(ck).IndexOf("已勾选", StringComparison.Ordinal) >= 0,
                    Dump(ck, "勾选未进树"));
            }

            if (string.IsNullOrEmpty(anchor) || string.IsNullOrEmpty(name))
            {
                run.CaseFail("bat-hash-not-stop", "缺页内锚点或姓名 ref");
            }
            else
            {
                ToolResult hash = s.InteractBatch(
                    id,
                    new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object> { ["action"] = "click", ["ref"] = anchor },
                        TypeRow(name, "锚点后仍填"),
                    });
                IntCases.Ok(
                    run,
                    "bat-hash-not-stop",
                    hash != null
                        && hash.Success
                        && StatusAt(hash, 1) == "success"
                        && IntCases.Tree(hash).IndexOf("锚点后仍填", StringComparison.Ordinal) >= 0,
                    Dump(hash, "只改 hash 不应停批"));
            }
        }

        private static void RunPartial(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = IntCases.OpenOk(run, s, http.Url("batch-form", "index.html"), "bat-mid-bad-ref");
            if (id == null)
            {
                IntCases.FailMany(
                    run,
                    "依赖 batch-form",
                    "bat-mid-bad-ref",
                    "bat-item-no-text",
                    "bat-item-bad-action",
                    "bat-item-click-no-ref",
                    "bat-item-press-no-key",
                    "bat-item-press-bad-key",
                    "bat-item-select-no-option",
                    "bat-item-select-bad-option",
                    "bat-item-download",
                    "bat-all-failed");
            }
            else
            {
                string tree = IntCases.WaitTree(s, id, true, "姓名", "邮箱");
                string name = IntCases.FindRef(tree, "姓名");
                string org = IntCases.FindRef(tree, "单位");
                string email = IntCases.FindRef(tree, "邮箱");
                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(org) || string.IsNullOrEmpty(email))
                {
                    IntCases.FailMany(
                        run,
                        "缺输入框 ref",
                        "bat-mid-bad-ref",
                        "bat-item-no-text",
                        "bat-item-bad-action",
                        "bat-item-click-no-ref",
                        "bat-item-press-no-key",
                        "bat-item-press-bad-key",
                        "bat-item-select-no-option",
                        "bat-item-select-bad-option",
                        "bat-item-download");
                }
                else
                {
                    ToolResult mid = s.InteractBatch(
                        id,
                        new List<Dictionary<string, object>>
                        {
                            TypeRow(name, "赵六"),
                            TypeRow("e9999", "不该写入"),
                            TypeRow(email, "z@z.com"),
                        });
                    string midTree = IntCases.Tree(mid);
                    bool midOk = PartialEnvelope(mid)
                        && StatusAt(mid, 0) == "success"
                        && StatusAt(mid, 1) == "failed"
                        && StatusAt(mid, 2) == "success"
                        && !string.IsNullOrEmpty(midTree)
                        && midTree.IndexOf("赵六", StringComparison.Ordinal) >= 0
                        && midTree.IndexOf("z@z.com", StringComparison.Ordinal) >= 0
                        && midTree.IndexOf("不该写入", StringComparison.Ordinal) < 0;
                    IntCases.Ok(run, "bat-mid-bad-ref", midOk, Dump(mid, "中间失败应跳过并继续"));

                    ToolResult noText = s.InteractBatch(
                        id,
                        new List<Dictionary<string, object>>
                        {
                            TypeRow(name, "钱七"),
                            new Dictionary<string, object> { ["action"] = "type", ["ref"] = org },
                            TypeRow(email, "q@q.com"),
                        });
                    IntCases.Ok(
                        run,
                        "bat-item-no-text",
                        PartialEnvelope(noText)
                            && StatusAt(noText, 0) == "success"
                            && StatusAt(noText, 1) == "failed"
                            && StatusAt(noText, 2) == "success",
                        Dump(noText, "缺 text 只该项失败"));

                    ToolResult badAct = s.InteractBatch(
                        id,
                        new List<Dictionary<string, object>>
                        {
                            TypeRow(name, "孙八"),
                            new Dictionary<string, object> { ["action"] = "foo", ["ref"] = org },
                            TypeRow(email, "s@s.com"),
                        });
                    IntCases.Ok(
                        run,
                        "bat-item-bad-action",
                        PartialEnvelope(badAct)
                            && StatusAt(badAct, 0) == "success"
                            && StatusAt(badAct, 1) == "failed"
                            && StatusAt(badAct, 2) == "success",
                        Dump(badAct, "假动词只该项失败"));

                    ToolResult noRef = s.InteractBatch(
                        id,
                        new List<Dictionary<string, object>>
                        {
                            TypeRow(name, "无ref前"),
                            new Dictionary<string, object> { ["action"] = "click" },
                            TypeRow(email, "w@w.com"),
                        });
                    IntCases.Ok(
                        run,
                        "bat-item-click-no-ref",
                        PartialEnvelope(noRef)
                            && StatusAt(noRef, 1) == "failed"
                            && StatusAt(noRef, 2) == "success",
                        Dump(noRef, "click 无 ref 只该项失败"));

                    ToolResult noKey = s.InteractBatch(
                        id,
                        new List<Dictionary<string, object>>
                        {
                            TypeRow(name, "无键前"),
                            new Dictionary<string, object> { ["action"] = "press", ["ref"] = name },
                            TypeRow(email, "k@k.com"),
                        });
                    IntCases.Ok(
                        run,
                        "bat-item-press-no-key",
                        PartialEnvelope(noKey)
                            && StatusAt(noKey, 1) == "failed"
                            && (ErrorAt(noKey, 1).IndexOf("key", StringComparison.OrdinalIgnoreCase) >= 0),
                        Dump(noKey, "press 无 key 只该项失败"));

                    ToolResult badKey = s.InteractBatch(
                        id,
                        new List<Dictionary<string, object>>
                        {
                            TypeRow(name, "坏键前"),
                            new Dictionary<string, object>
                            {
                                ["action"] = "press",
                                ["ref"] = name,
                                ["key"] = "Control",
                            },
                            TypeRow(email, "c@c.com"),
                        });
                    IntCases.Ok(
                        run,
                        "bat-item-press-bad-key",
                        PartialEnvelope(badKey)
                            && StatusAt(badKey, 1) == "failed"
                            && IntCases.ContainsAny(ErrorAt(badKey, 1), "白名单", "key"),
                        Dump(badKey, "非法 key 只该项失败"));

                    string kind = IntCases.FindRef(tree, "课题类型");
                    if (string.IsNullOrEmpty(kind))
                    {
                        IntCases.FailMany(run, "缺课题类型 ref", "bat-item-select-no-option", "bat-item-select-bad-option");
                    }
                    else
                    {
                        ToolResult noOpt = s.InteractBatch(
                            id,
                            new List<Dictionary<string, object>>
                            {
                                TypeRow(name, "无选项前"),
                                new Dictionary<string, object> { ["action"] = "select", ["ref"] = kind },
                                TypeRow(email, "o@o.com"),
                            });
                        IntCases.Ok(
                            run,
                            "bat-item-select-no-option",
                            PartialEnvelope(noOpt)
                                && StatusAt(noOpt, 1) == "failed"
                                && ErrorAt(noOpt, 1).IndexOf("option", StringComparison.OrdinalIgnoreCase) >= 0,
                            Dump(noOpt, "select 无 option 只该项失败"));

                        ToolResult badOpt = s.InteractBatch(
                            id,
                            new List<Dictionary<string, object>>
                            {
                                TypeRow(name, "坏选项前"),
                                new Dictionary<string, object>
                                {
                                    ["action"] = "select",
                                    ["ref"] = kind,
                                    ["option"] = "火星课题",
                                },
                                TypeRow(email, "m@m.com"),
                            });
                        IntCases.Ok(
                            run,
                            "bat-item-select-bad-option",
                            PartialEnvelope(badOpt)
                                && StatusAt(badOpt, 1) == "failed"
                                && StatusAt(badOpt, 2) == "success",
                            Dump(badOpt, "假 option 只该项失败"));
                    }

                    ToolResult dl = s.InteractBatch(
                        id,
                        new List<Dictionary<string, object>>
                        {
                            TypeRow(name, "下载前"),
                            new Dictionary<string, object> { ["action"] = "download", ["ref"] = email },
                            TypeRow(email, "d@d.com"),
                        });
                    IntCases.Ok(
                        run,
                        "bat-item-download",
                        PartialEnvelope(dl)
                            && StatusAt(dl, 1) == "failed"
                            && StatusAt(dl, 2) == "success",
                        Dump(dl, "download 不能进 actions"));
                }

                ToolResult allFail = s.InteractBatch(
                    id,
                    new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object> { ["action"] = "click", ["ref"] = "e8888" },
                        new Dictionary<string, object> { ["action"] = "click", ["ref"] = "e9999" },
                    });
                bool allFailOk = allFail != null
                    && !allFail.Success
                    && Outcome(allFail) == "all_failed"
                    && FirstLine(allFail).StartsWith("全部失败", StringComparison.Ordinal)
                    && StatusAt(allFail, 0) == "failed"
                    && StatusAt(allFail, 1) == "failed";
                IntCases.Ok(run, "bat-all-failed", allFailOk, Dump(allFail, "应全部失败"));
            }

            string d18 = IntCases.OpenOk(run, s, http.Url("d18-risk", "index.html"), "bat-d18-skip");
            if (d18 == null)
            {
                run.CaseFail("bat-d18-login-continue", "依赖 d18-risk");
            }
            else
            {
            string dtree = IntCases.WaitTree(s, d18, true, "搜索框", "密码", "提交");
            string q = IntCases.FindRef(dtree, "搜索框");
            string pwd = IntCases.FindRef(dtree, "密码");
            string submit = IntCases.FindRef(dtree, "提交", "button");
            if (string.IsNullOrEmpty(q) || string.IsNullOrEmpty(pwd) || string.IsNullOrEmpty(submit))
            {
                run.CaseFail("bat-d18-skip", "缺搜索/密码/提交 ref");
                run.CaseFail("bat-d18-login-continue", "缺搜索/密码/提交 ref");
            }
            else
            {

            ToolResult d18r = s.InteractBatch(
                d18,
                new List<Dictionary<string, object>>
                {
                    TypeRow(q, "检索词"),
                    TypeRow(pwd, "secret"),
                    new Dictionary<string, object> { ["action"] = "click", ["ref"] = submit },
                });
            string dsnap = IntCases.Tree(d18r);
            bool d18ok = PartialEnvelope(d18r)
                && StatusAt(d18r, 0) == "success"
                && StatusAt(d18r, 1) == "failed"
                && StatusAt(d18r, 2) == "success"
                && dsnap.IndexOf("检索词", StringComparison.Ordinal) >= 0
                && dsnap.IndexOf("提交已点", StringComparison.Ordinal) >= 0;
            IntCases.Ok(run, "bat-d18-skip", d18ok, Dump(d18r, "D18 应只拦密码项"));

            string loginBtn = IntCases.FindRef(dtree, "登录");
            if (string.IsNullOrEmpty(q) || string.IsNullOrEmpty(loginBtn) || string.IsNullOrEmpty(submit))
            {
                run.CaseFail("bat-d18-login-continue", "缺搜索/登录/提交 ref");
            }
            else
            {
                string d18b = IntCases.OpenOk(run, s, http.Url("d18-risk", "index.html"), "bat-d18-login-continue");
                if (d18b != null)
                {
                    string t2 = IntCases.WaitTree(s, d18b, true, "搜索框", "登录", "提交");
                    q = IntCases.FindRef(t2, "搜索框") ?? q;
                    loginBtn = IntCases.FindRef(t2, "登录") ?? loginBtn;
                    submit = IntCases.FindRef(t2, "提交", "button") ?? submit;
                    ToolResult login = s.InteractBatch(
                        d18b,
                        new List<Dictionary<string, object>>
                        {
                            TypeRow(q, "可搜"),
                            new Dictionary<string, object> { ["action"] = "click", ["ref"] = loginBtn },
                            new Dictionary<string, object> { ["action"] = "click", ["ref"] = submit },
                        });
                    string lt = IntCases.Tree(login);
                    IntCases.Ok(
                        run,
                        "bat-d18-login-continue",
                        PartialEnvelope(login)
                            && StatusAt(login, 1) == "failed"
                            && StatusAt(login, 2) == "success"
                            && lt.IndexOf("提交已点", StringComparison.Ordinal) >= 0
                            && lt.IndexOf("登录已点", StringComparison.Ordinal) < 0,
                        Dump(login, "登录应拦、提交应做成"));
                }
            }
            }
            }

            string dis = IntCases.OpenOk(run, s, http.Url("disabled-input", "index.html"), "bat-type-disabled-continue");
            if (dis == null)
            {
                return;
            }

            string dt = IntCases.WaitTree(s, dis, true, "禁用框", "可填框");
            string bad = IntCases.FindRef(dt, "禁用框");
            string okBox = IntCases.FindRef(dt, "可填框");
            if (string.IsNullOrEmpty(bad) || string.IsNullOrEmpty(okBox))
            {
                run.CaseFail("bat-type-disabled-continue", "缺禁用框或可填框 ref");
                return;
            }

            ToolResult disr = s.InteractBatch(
                dis,
                new List<Dictionary<string, object>>
                {
                    TypeRow(bad, "改"),
                    TypeRow(okBox, "已填"),
                });
            string ds = IntCases.Tree(disr);
            IntCases.Ok(
                run,
                "bat-type-disabled-continue",
                PartialEnvelope(disr)
                    && StatusAt(disr, 0) == "failed"
                    && StatusAt(disr, 1) == "success"
                    && ds.IndexOf("已填", StringComparison.Ordinal) >= 0
                    && ds.IndexOf("不该改到", StringComparison.Ordinal) < 0,
                Dump(disr, "禁用应失败并继续填可填框"));
        }

        private static void RunStop(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = IntCases.OpenOk(run, s, http.Url("batch-form", "index.html"), "bat-stop-confirm");
            if (id == null)
            {
                IntCases.FailMany(
                    run,
                    "依赖 batch-form",
                    "bat-stop-confirm",
                    "bat-stop-nav",
                    "bat-stop-newtab");
            }
            else
            {
                string tree = IntCases.WaitTree(s, id, true, "弹出确认", "姓名", "去完成页", "新窗口");
                string confirm = IntCases.FindRef(tree, "弹出确认");
                string name = IntCases.FindRef(tree, "姓名");
                string go = IntCases.FindRef(tree, "去完成页");
                string tab = IntCases.FindRef(tree, "新窗口");
                if (string.IsNullOrEmpty(confirm) || string.IsNullOrEmpty(name))
                {
                    run.CaseFail("bat-stop-confirm", "缺弹出确认或姓名 ref");
                }
                else
                {
                    ToolResult dlg = s.InteractBatch(
                        id,
                        new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object> { ["action"] = "click", ["ref"] = confirm },
                            TypeRow(name, "不应填入"),
                        });
                    string dlgTree = IntCases.Tree(dlg);
                    bool dlgOk = PartialEnvelope(dlg)
                        && StatusAt(dlg, 0) == "success"
                        && StatusAt(dlg, 1) == "skipped"
                        && dlgTree.IndexOf("不应填入", StringComparison.Ordinal) < 0;
                    IntCases.Ok(run, "bat-stop-confirm", dlgOk, Dump(dlg, "确认框后应停批"));
                    string cancel = IntCases.FindRef(dlgTree, "取消") ?? IntCases.FindRef(dlgTree, "确定");
                    if (!string.IsNullOrEmpty(cancel))
                    {
                        s.Interact(id, "click", cancel);
                    }
                }

                tree = IntCases.WaitTree(s, id, false, "去完成页");
                go = IntCases.FindRef(tree, "去完成页") ?? go;
                name = IntCases.FindRef(tree, "姓名") ?? name;
                if (string.IsNullOrEmpty(go) || string.IsNullOrEmpty(name))
                {
                    run.CaseFail("bat-stop-nav", "缺去完成页或姓名 ref");
                }
                else
                {
                    ToolResult nav = s.InteractBatch(
                        id,
                        new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object> { ["action"] = "click", ["ref"] = go },
                            TypeRow(name, "跳转后不填"),
                        });
                    string url = IntCases.Str(nav, "url");
                    bool navOk = PartialEnvelope(nav)
                        && StatusAt(nav, 0) == "success"
                        && StatusAt(nav, 1) == "skipped"
                        && url.IndexOf("done.html", StringComparison.OrdinalIgnoreCase) >= 0;
                    IntCases.Ok(run, "bat-stop-nav", navOk, Dump(nav, "跳转后应停批"));
                }

                string tabId = IntCases.OpenOk(run, s, http.Url("batch-form", "index.html"), "bat-stop-newtab");
                if (tabId == null)
                {
                    return;
                }

                string ttree = IntCases.WaitTree(s, tabId, true, "新窗口", "姓名");
                tab = IntCases.FindRef(ttree, "新窗口");
                name = IntCases.FindRef(ttree, "姓名");
                if (string.IsNullOrEmpty(tab) || string.IsNullOrEmpty(name))
                {
                    run.CaseFail("bat-stop-newtab", "缺新窗口或姓名 ref");
                }
                else
                {
                    ToolResult nt = s.InteractBatch(
                        tabId,
                        new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object> { ["action"] = "click", ["ref"] = tab },
                            TypeRow(name, "新标签不填"),
                        });
                    bool ntOk = PartialEnvelope(nt)
                        && StatusAt(nt, 0) == "success"
                        && StatusAt(nt, 1) == "skipped";
                    IntCases.Ok(run, "bat-stop-newtab", ntOk, Dump(nt, "新标签后应停批"));
                }
            }

            string ifr = IntCases.OpenOk(run, s, http.Url("batch-form", "outer.html"), "bat-stop-iframe");
            if (ifr == null)
            {
                return;
            }

            string ot = IntCases.WaitTree(s, ifr, true, "教师工作台", "外壳姓名");
            string frame = IntCases.FindRef(ot, "教师工作台") ?? IntCases.FindIframeRef(ot);
            string outer = IntCases.FindRef(ot, "外壳姓名");
            if (string.IsNullOrEmpty(frame) || string.IsNullOrEmpty(outer))
            {
                run.CaseFail("bat-stop-iframe", "缺 iframe 行或外壳姓名 ref");
                return;
            }

            ToolResult exp = s.InteractBatch(
                ifr,
                new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object> { ["action"] = "click", ["ref"] = frame },
                    TypeRow(outer, "外壳不应填"),
                });
            string et = IntCases.Tree(exp);
            bool expOk = PartialEnvelope(exp)
                && StatusAt(exp, 0) == "success"
                && StatusAt(exp, 1) == "skipped"
                && (et.IndexOf("课程名", StringComparison.Ordinal) >= 0
                    || IntCases.Flag(exp, "expanded_frame"))
                && et.IndexOf("外壳不应填", StringComparison.Ordinal) < 0;
            IntCases.Ok(run, "bat-stop-iframe", expOk, Dump(exp, "展开 iframe 后应停批"));
        }

        private static Dictionary<string, object> TypeRow(string refId, string text)
        {
            return new Dictionary<string, object>
            {
                ["action"] = "type",
                ["ref"] = refId ?? "",
                ["text"] = text ?? "",
            };
        }

        private static bool PartialEnvelope(ToolResult r)
        {
            return r != null
                && !r.Success
                && Outcome(r) == "partial_succeeded"
                && FirstLine(r).StartsWith("部分成功", StringComparison.Ordinal);
        }

        private static string Outcome(ToolResult r)
        {
            return IntCases.Str(r, "batch_outcome");
        }

        private static string FirstLine(ToolResult r)
        {
            string msg = IntCases.Str(r, "message");
            if (string.IsNullOrEmpty(msg) && r != null)
            {
                msg = r.Error ?? "";
            }

            int nl = msg.IndexOf('\n');
            return nl < 0 ? msg : msg.Substring(0, nl);
        }

        private static bool AllStatus(ToolResult r, string status)
        {
            List<Dictionary<string, object>> rows = Rows(r);
            if (rows.Count == 0)
            {
                return false;
            }

            for (int i = 0; i < rows.Count; i++)
            {
                if (!string.Equals(Cell(rows[i], "status"), status, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }

        private static string ErrorAt(ToolResult r, int index)
        {
            List<Dictionary<string, object>> rows = Rows(r);
            if (index < 0 || index >= rows.Count)
            {
                return "";
            }

            return Cell(rows[index], "error");
        }

        private static int IntAt(ToolResult r, int index, string key)
        {
            List<Dictionary<string, object>> rows = Rows(r);
            if (index < 0 || index >= rows.Count)
            {
                return -1;
            }

            string raw = Cell(rows[index], key);
            int n;
            return int.TryParse(raw, out n) ? n : -1;
        }

        private static string StatusAt(ToolResult r, int index)
        {
            List<Dictionary<string, object>> rows = Rows(r);
            if (index < 0 || index >= rows.Count)
            {
                return "";
            }

            return Cell(rows[index], "status");
        }

        private static List<Dictionary<string, object>> Rows(ToolResult r)
        {
            var list = new List<Dictionary<string, object>>();
            Dictionary<string, object> data = IntCases.Data(r);
            if (data == null || !data.ContainsKey("results") || data["results"] == null)
            {
                return list;
            }

            object raw = data["results"];
            var ja = raw as JArray;
            IEnumerable seq = ja != null ? (IEnumerable)ja : raw as IEnumerable;
            if (seq == null || raw is string)
            {
                return list;
            }

            foreach (object item in seq)
            {
                Dictionary<string, object> row = Coerce(item);
                if (row != null)
                {
                    list.Add(row);
                }
            }

            return list;
        }

        private static Dictionary<string, object> Coerce(object item)
        {
            if (item is Dictionary<string, object> dict)
            {
                return dict;
            }

            var jo = item as JObject;
            return jo != null ? jo.ToObject<Dictionary<string, object>>() : null;
        }

        private static string Cell(Dictionary<string, object> row, string key)
        {
            if (row == null || !row.ContainsKey(key) || row[key] == null)
            {
                return "";
            }

            return Convert.ToString(row[key]) ?? "";
        }

        private static bool ContainsAll(string text, params string[] parts)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            for (int i = 0; i < parts.Length; i++)
            {
                if (text.IndexOf(parts[i], StringComparison.Ordinal) < 0)
                {
                    return false;
                }
            }

            return true;
        }

        private static string Dump(ToolResult r, string fallback)
        {
            if (r == null)
            {
                return fallback;
            }

            string extra = Outcome(r);
            if (!string.IsNullOrEmpty(r.Error))
            {
                return fallback + " · " + extra + " · " + r.Error;
            }

            string msg = IntCases.Str(r, "message");
            return string.IsNullOrEmpty(msg) ? fallback + " · " + extra : fallback + " · " + extra + " · " + msg;
        }
    }
}
