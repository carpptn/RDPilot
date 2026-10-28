
namespace RDPilot.Control;


            internal sealed class ControlContextChain
            {
                readonly string chainId;
                readonly int fallbackLimit;
                int fallbackCount;
                int compactionCount;

                internal ControlContextChain(
                    string commandId,
                    bool enabled,
                    bool compactionEnabled,
                    int fallbackLimit)
                {
                    chainId = commandId.Length <= 8 ? commandId : commandId[..8];
                    Enabled = enabled;
                    CompactionEnabled = enabled && compactionEnabled;
                    this.fallbackLimit = Math.Max(1, fallbackLimit);
                }

                internal bool Enabled { get; private set; }
                internal bool CompactionEnabled { get; private set; }
                internal string? PreviousResponseId { get; private set; }
                internal string? PreviousResponseIdForRequest =>
                    Enabled ? PreviousResponseId : null;
                internal int TurnCount { get; private set; }
                internal int RestartCount { get; private set; }

                internal void LogStart(string model)
                {
                    if (!Enabled)
                    {
                        Console.WriteLine("[context] control chaining disabled; explicit application history will be used.");
                        return;
                    }

                    var effectiveContext = SupportsReasoningContext(model)
                        ? Options.ControlReasoningContext
                        : "model_default";
                    var compaction = CompactionEnabled
                        ? Options.ControlContextCompactThreshold.ToString(CultureInfo.InvariantCulture)
                        : "off";
                    Console.WriteLine(
                        $"[context] started control chain id={chainId}; reasoning_context={effectiveContext}; compact_threshold={compaction}.");
                }

                internal void LogRequest(int step)
                {
                    if (!Enabled)
                        return;

                    if (PreviousResponseId is null)
                    {
                        Console.WriteLine($"[context] chain={chainId}; step={step}; starting API turn without previous_response_id.");
                        return;
                    }

                    Console.WriteLine(
                        $"[context] chain={chainId}; step={step}; continuing turn={TurnCount + 1}; previous={ShortResponseId(PreviousResponseId)}.");
                }

                internal void RecordResult(
                    string? completedResponseId,
                    bool contextFallbackUsed,
                    bool compactionFallbackUsed,
                    bool compactionOccurred)
                {
                    if (compactionFallbackUsed && CompactionEnabled)
                    {
                        CompactionEnabled = false;
                        Console.WriteLine(
                            $"[context] chain={chainId}; server-side compaction disabled for the remainder of this task.");
                    }

                    if (contextFallbackUsed)
                    {
                        PreviousResponseId = null;
                        fallbackCount++;
                        RestartCount++;
                        CurrentRun.RunControlContextRestarts++;
                        Console.WriteLine(
                            $"[context] chain={chainId}; restarted from checkpoint; fallback={fallbackCount}/{fallbackLimit}.");
                        if (fallbackCount >= fallbackLimit)
                        {
                            Enabled = false;
                            CompactionEnabled = false;
                            Console.WriteLine(
                                $"[context] chain={chainId}; disabled after {fallbackCount} state failures; continuing with explicit application history.");
                        }
                    }

                    if (compactionOccurred)
                    {
                        compactionCount++;
                        Console.WriteLine(
                            $"[context] chain={chainId}; server-side compaction completed; count={compactionCount}.");
                    }

                    if (!Enabled || string.IsNullOrWhiteSpace(completedResponseId))
                        return;

                    PreviousResponseId = completedResponseId;
                    TurnCount++;
                    CurrentRun.RunControlContextTurns++;
                    Console.WriteLine(
                        $"[context] chain={chainId}; finalized turn={TurnCount}; response={ShortResponseId(completedResponseId)}.");
                }

                internal void LogClose(ControlRunOutcome outcome)
                {
                    if (!Enabled && TurnCount == 0 && RestartCount == 0)
                        return;

                    Console.WriteLine(
                        $"[context] closed chain id={chainId}; outcome={outcome}; turns={TurnCount}; restarts={RestartCount}; compactions={compactionCount}.");
                }

                static string ShortResponseId(string responseId) =>
                    responseId.Length <= 18 ? responseId : responseId[..18] + "...";
            }
