using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace WordAddIn1
{
    public static class CheckpointHistoryStore
    {
        private static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            NullValueHandling = NullValueHandling.Ignore
        };

        public static DocumentCheckpointSession TryLoad(string docId)
        {
            string path = HistoryPath(docId);
            if (!File.Exists(path))
            {
                return null;
            }

            try
            {
                string json = File.ReadAllText(path);
                HistoryFile file = JsonConvert.DeserializeObject<HistoryFile>(json, Settings);
                return file == null ? null : ToSession(docId, file);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[Checkpoint] history load failed: " + ex.Message);
                return null;
            }
        }

        public static void MarkOrdinaryRetired(string historyPath)
        {
            if (string.IsNullOrEmpty(historyPath) || !File.Exists(historyPath))
            {
                return;
            }

            try
            {
                string json = File.ReadAllText(historyPath);
                HistoryFile file = JsonConvert.DeserializeObject<HistoryFile>(json, Settings);
                if (file == null || file.Entries == null)
                {
                    return;
                }

                bool changed = false;
                foreach (HistoryEntry entry in file.Entries)
                {
                    if (entry == null
                        || !string.Equals(entry.Kind, "checkpoint", StringComparison.OrdinalIgnoreCase)
                        || entry.UserRollback
                        || entry.Retired)
                    {
                        continue;
                    }

                    entry.Retired = true;
                    changed = true;
                }

                if (!changed)
                {
                    return;
                }

                string temp = historyPath + ".tmp";
                File.WriteAllText(temp, JsonConvert.SerializeObject(file, Formatting.Indented, Settings));
                File.Replace(temp, historyPath, null);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[Checkpoint] retire history failed: " + ex.Message);
            }
        }

        public static void Save(DocumentCheckpointSession session)
        {
            if (session == null || string.IsNullOrEmpty(session.DocUuid))
            {
                return;
            }

            string dir = DocumentCheckpointService.GetDocCheckpointDirectory(session.DocUuid);
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "history.json");
            string temp = path + ".tmp";
            string json = JsonConvert.SerializeObject(FromSession(session), Formatting.Indented, Settings);
            File.WriteAllText(temp, json);
            if (File.Exists(path))
            {
                File.Replace(temp, path, null);
            }
            else
            {
                File.Move(temp, path);
            }
        }

        private static string HistoryPath(string docId)
        {
            return Path.Combine(DocumentCheckpointService.GetDocCheckpointDirectory(docId), "history.json");
        }

        private static DocumentCheckpointSession ToSession(string docId, HistoryFile file)
        {
            var session = new DocumentCheckpointSession
            {
                DocUuid = string.IsNullOrEmpty(file.DocId) ? docId : file.DocId,
                Name = file.Name,
                Path = file.Path,
                Kind = file.Kind,
                CurrentBranch = file.CurrentBranch > 0 ? file.CurrentBranch : 1,
                PositionBranch = file.Position == null ? 0 : file.Position.Branch,
                PositionTimestamp = file.Position?.Timestamp
            };

            if (file.Branches != null)
            {
                foreach (HistoryBranch branch in file.Branches)
                {
                    if (branch == null || branch.Id <= 0)
                    {
                        continue;
                    }

                    session.Branches.Add(new CheckpointBranch
                    {
                        Id = branch.Id,
                        ForkedFromBranch = branch.ForkedFrom?.Branch ?? 0,
                        ForkedFromTimestamp = branch.ForkedFrom?.Timestamp
                    });
                }
            }

            if (session.Branches.Count == 0)
            {
                session.Branches.Add(new CheckpointBranch { Id = 1 });
            }

            if (file.Entries == null)
            {
                return session;
            }

            foreach (HistoryEntry entry in file.Entries)
            {
                if (entry == null || string.IsNullOrEmpty(entry.Kind))
                {
                    continue;
                }

                bool checkpoint = string.Equals(entry.Kind, "checkpoint", StringComparison.OrdinalIgnoreCase);
                var log = new OperationLogEntry
                {
                    Kind = checkpoint ? OperationLogEntryKind.Checkpoint : OperationLogEntryKind.Action,
                    Branch = entry.Branch > 0 ? entry.Branch : 1,
                    Timestamp = entry.Timestamp,
                    UserRollback = entry.UserRollback,
                    Hash = entry.Hash ?? "",
                    Retired = entry.Retired,
                    ToolName = entry.Tool,
                    ActionLineBody = entry.Arguments
                };
                if (entry.Operations != null)
                {
                    foreach (HistoryOperation op in entry.Operations)
                    {
                        if (op == null)
                        {
                            continue;
                        }

                        log.Operations.Add(new CheckpointOperationRecord
                        {
                            Tool = op.Tool,
                            Arguments = op.Arguments
                        });
                    }
                }

                session.Entries.Add(log);
                if (!checkpoint || string.IsNullOrEmpty(entry.Timestamp))
                {
                    continue;
                }

                session.CheckpointTimestampsInOrder.Add(entry.Timestamp);
                if (entry.Caret != null)
                {
                    session.CaretByTimestamp[entry.Timestamp] = entry.Caret;
                }
            }

            return session;
        }

        private static HistoryFile FromSession(DocumentCheckpointSession session)
        {
            var file = new HistoryFile
            {
                DocId = session.DocUuid,
                Name = session.Name,
                Path = session.Path,
                Kind = session.Kind,
                CurrentBranch = session.CurrentBranch > 0 ? session.CurrentBranch : 1,
                Branches = new List<HistoryBranch>(),
                Entries = new List<HistoryEntry>()
            };
            if (session.PositionBranch > 0 && !string.IsNullOrEmpty(session.PositionTimestamp))
            {
                file.Position = new HistoryPosition
                {
                    Branch = session.PositionBranch,
                    Timestamp = session.PositionTimestamp
                };
            }

            foreach (CheckpointBranch branch in session.Branches)
            {
                var row = new HistoryBranch { Id = branch.Id };
                if (branch.ForkedFromBranch > 0 && !string.IsNullOrEmpty(branch.ForkedFromTimestamp))
                {
                    row.ForkedFrom = new HistoryFork
                    {
                        Branch = branch.ForkedFromBranch,
                        Timestamp = branch.ForkedFromTimestamp
                    };
                }

                file.Branches.Add(row);
            }

            foreach (OperationLogEntry entry in session.Entries)
            {
                bool checkpoint = entry.Kind == OperationLogEntryKind.Checkpoint;
                var row = new HistoryEntry
                {
                    Kind = checkpoint ? "checkpoint" : "action",
                    Branch = entry.Branch > 0 ? entry.Branch : 1,
                    Timestamp = entry.Timestamp,
                    UserRollback = entry.UserRollback,
                    Hash = string.IsNullOrEmpty(entry.Hash) ? null : entry.Hash,
                    Retired = entry.Retired,
                    Tool = entry.ToolName,
                    Arguments = entry.ActionLineBody
                };
                if (checkpoint && !string.IsNullOrEmpty(entry.Timestamp)
                    && session.CaretByTimestamp.TryGetValue(entry.Timestamp, out CheckpointCaret caret))
                {
                    row.Caret = caret;
                }

                if (entry.UserRollback && entry.Operations.Count > 0)
                {
                    row.Operations = new List<HistoryOperation>();
                    foreach (CheckpointOperationRecord op in entry.Operations)
                    {
                        row.Operations.Add(new HistoryOperation
                        {
                            Tool = op.Tool,
                            Arguments = op.Arguments
                        });
                    }
                }

                file.Entries.Add(row);
            }

            return file;
        }

        private sealed class HistoryFile
        {
            public string DocId { get; set; }
            public string Name { get; set; }
            public string Path { get; set; }
            public string Kind { get; set; }
            public int CurrentBranch { get; set; }
            public HistoryPosition Position { get; set; }
            public List<HistoryBranch> Branches { get; set; }
            public List<HistoryEntry> Entries { get; set; }
        }

        private sealed class HistoryPosition
        {
            public int Branch { get; set; }
            public string Timestamp { get; set; }
        }

        private sealed class HistoryBranch
        {
            public int Id { get; set; }
            public HistoryFork ForkedFrom { get; set; }
        }

        private sealed class HistoryFork
        {
            public int Branch { get; set; }
            public string Timestamp { get; set; }
        }

        private sealed class HistoryEntry
        {
            public string Kind { get; set; }
            public int Branch { get; set; }
            public string Timestamp { get; set; }
            public bool UserRollback { get; set; }
            public string Hash { get; set; }
            public bool Retired { get; set; }
            public CheckpointCaret Caret { get; set; }
            public string Tool { get; set; }
            public string Arguments { get; set; }
            public List<HistoryOperation> Operations { get; set; }
        }

        private sealed class HistoryOperation
        {
            public string Tool { get; set; }
            public string Arguments { get; set; }
        }
    }
}
