using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>附着的 Chrome / Edge 上挂起的脚本确认框。框仍是浏览器原来的那一扇。</summary>
    internal static class BrowserScriptDialogAttach
    {
        private static readonly object Gate = new object();
        private static readonly Dictionary<string, BrowserScriptDialogSession> Sessions =
            new Dictionary<string, BrowserScriptDialogSession>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, TaskCompletionSource<bool>> Watches =
            new Dictionary<string, TaskCompletionSource<bool>>(StringComparer.OrdinalIgnoreCase);

        public static Task Arm(string tabUuid)
        {
            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            lock (Gate)
            {
                Watches[tabUuid ?? ""] = tcs;
                BrowserScriptDialogSession existing;
                if (Sessions.TryGetValue(tabUuid ?? "", out existing) && existing != null && existing.IsPending)
                {
                    tcs.TrySetResult(true);
                }
            }

            return tcs.Task;
        }

        public static BrowserScriptDialogSession Peek(string tabUuid)
        {
            lock (Gate)
            {
                BrowserScriptDialogSession existing;
                if (Sessions.TryGetValue(tabUuid ?? "", out existing) && existing != null && existing.IsPending)
                {
                    return existing;
                }
            }

            return null;
        }

        public static async Task<BrowserScriptDialogSession> SyncAsync(BrowserChannel channel)
        {
            if (channel == null || string.IsNullOrWhiteSpace(channel.TabUuid))
            {
                return null;
            }

            JObject result;
            try
            {
                var req = new JObject
                {
                    ["tab_uuid"] = channel.TabUuid,
                    ["op"] = "get"
                };
                result = await ExtensionHost
                    .RpcAsync(channel.TabUuid, "rpc.script_dialog", req, TimeSpan.FromSeconds(8))
                    .ConfigureAwait(true);
            }
            catch
            {
                return Peek(channel.TabUuid);
            }

            JToken dialog = result != null ? result["dialog"] : null;
            if (dialog == null || dialog.Type == JTokenType.Null)
            {
                Clear(channel.TabUuid);
                return null;
            }

            return Remember(channel.TabUuid, dialog);
        }

        public static BrowserScriptDialogSession Remember(string tabUuid, JToken dialog)
        {
            if (string.IsNullOrWhiteSpace(tabUuid) || dialog == null || dialog.Type == JTokenType.Null)
            {
                return null;
            }

            string type = (string)dialog["type"] ?? "confirm";
            string message = (string)dialog["message"] ?? "";
            string prompt = (string)dialog["defaultPrompt"] ?? "";
            BrowserScriptDialogKind kind = MapKind(type);
            lock (Gate)
            {
                BrowserScriptDialogSession existing;
                if (Sessions.TryGetValue(tabUuid, out existing)
                    && existing != null
                    && existing.IsPending
                    && existing.Kind == kind
                    && string.Equals(existing.Message, message, StringComparison.Ordinal))
                {
                    SignalLocked(tabUuid);
                    return existing;
                }

                BrowserScriptDialogSession created = BrowserScriptDialogSession.Create(
                    kind,
                    message,
                    prompt,
                    _ => { },
                    () => { });
                Sessions[tabUuid] = created;
                SignalLocked(tabUuid);
                return created;
            }
        }

        public static BrowserSnapshotResult Publish(BrowserChannel channel, BrowserScriptDialogSession session)
        {
            BrowserScriptDialogView view = BrowserScriptDialog.Render(session);
            BrowserRefStore.Replace(channel.ChannelId, view.Refs);
            return BrowserSnapshotResult.Ok(
                channel,
                channel.Url,
                channel.Title,
                "overview",
                null,
                view.TreeText,
                false,
                null,
                null);
        }

        public static async Task<BrowserInteractResult> InteractAsync(
            BrowserChannel channel,
            BrowserScriptDialogSession session,
            string action,
            string refId,
            string text)
        {
            BrowserScriptDialogStep step = BrowserScriptDialog.Apply(
                session,
                channel.ChannelId,
                action,
                refId,
                text,
                out string error);
            if (step == BrowserScriptDialogStep.Reject || step == BrowserScriptDialogStep.NotPending)
            {
                return BrowserInteractResult.Fail(error ?? BrowserScriptDialog.MustHandleFirst);
            }

            if (step == BrowserScriptDialogStep.Accept || step == BrowserScriptDialogStep.Dismiss)
            {
                var req = new JObject
                {
                    ["tab_uuid"] = channel.TabUuid,
                    ["op"] = step == BrowserScriptDialogStep.Accept ? "accept" : "dismiss",
                    ["prompt_text"] = session.Text ?? ""
                };
                try
                {
                    await ExtensionHost
                        .RpcAsync(channel.TabUuid, "rpc.script_dialog", req, TimeSpan.FromSeconds(15))
                        .ConfigureAwait(true);
                }
                catch (Exception ex)
                {
                    return BrowserInteractResult.Fail(ex.Message ?? "确认框操作失败");
                }

                BrowserScriptDialog.Commit(session, step);
                Clear(channel.TabUuid);
            }

            return await BrowserResnapshotAfterAction
                .CompleteInteractAsync(
                    channel,
                    action,
                    string.IsNullOrWhiteSpace(refId) ? null : refId.Trim(),
                    BrowserResnapshotLayer.Shell())
                .ConfigureAwait(true);
        }

        public static void Clear(string tabUuid)
        {
            lock (Gate)
            {
                Sessions.Remove(tabUuid ?? "");
            }
        }

        private static void SignalLocked(string tabUuid)
        {
            TaskCompletionSource<bool> watch;
            if (Watches.TryGetValue(tabUuid ?? "", out watch) && watch != null)
            {
                watch.TrySetResult(true);
            }
        }

        private static BrowserScriptDialogKind MapKind(string type)
        {
            if (string.Equals(type, "alert", StringComparison.OrdinalIgnoreCase))
            {
                return BrowserScriptDialogKind.Alert;
            }

            if (string.Equals(type, "prompt", StringComparison.OrdinalIgnoreCase))
            {
                return BrowserScriptDialogKind.Prompt;
            }

            if (string.Equals(type, "beforeunload", StringComparison.OrdinalIgnoreCase))
            {
                return BrowserScriptDialogKind.Beforeunload;
            }

            return BrowserScriptDialogKind.Confirm;
        }
    }
}
