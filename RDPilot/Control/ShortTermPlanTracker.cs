
namespace RDPilot.Control;


            internal sealed class ShortTermPlanTracker
            {
                string? plan;
                string status = "none";
                string? revisionReason;
                string[] plannedInputs = [];
                string? waypoint;
                string? originStateId;
                string? expectedStateId;
                double? planConfidence;
                int currentInputIndex;

                internal bool HasActivePlan =>
                    plan is not null && status == "active";

                internal bool HasExecutableDirectionalPlan =>
                    HasActivePlan && currentInputIndex < plannedInputs.Length;

                internal string Status => status;

                internal int RemainingInputCount =>
                    Math.Max(0, plannedInputs.Length - currentInputIndex);

                internal double EffectiveConfidence => planConfidence ?? 0;

                internal IReadOnlyList<string> RemainingInputs(string? currentStateId) =>
                    HasExecutableDirectionalPlan &&
                    (expectedStateId is null || currentStateId is null || expectedStateId == currentStateId)
                        ? plannedInputs.Skip(currentInputIndex).ToArray() : [];

                internal IReadOnlyList<string> PrepareInputsForProposal(ActionDto action, string? currentStateId)
                {
                    Update(action, currentStateId);
                    // Older action payloads supplied an inline route without persistent plan metadata.
                    return string.IsNullOrWhiteSpace(action.PlanStatus) && !HasActivePlan
                        ? action.PlannedInputs ?? [] : RemainingInputs(currentStateId);
                }

                internal void Update(ActionDto action, string? currentStateId = null)
                {
                    var requestedStatus = action.PlanStatus?.Trim().ToLowerInvariant();
                    if (string.IsNullOrWhiteSpace(requestedStatus))
                        return;

                    if (requestedStatus == "none")
                    {
                        if (plan is not null || status != "none")
                            Console.WriteLine("[plan] cleared.");
                        plan = null;
                        status = "none";
                        revisionReason = null;
                        plannedInputs = [];
                        waypoint = null;
                        originStateId = null;
                        expectedStateId = null;
                        planConfidence = null;
                        currentInputIndex = 0;
                        return;
                    }

                    var nextPlan = NormalizePlanText(action.ShortTermPlan, 420);
                    if (nextPlan is null)
                    {
                        Invalidate("the model marked a plan active without supplying its steps");
                        return;
                    }

                    var nextReason = NormalizePlanText(action.PlanRevisionReason, 200);
                    var nextInputs = NormalizePlannedInputs(action.PlannedInputs);
                    if (BatchPlanner.HasImmediateDirectionalReversal(nextInputs))
                    {
                        Invalidate(
                            "the proposed route contains an immediate opposing-direction reversal and is not a coherent path to one waypoint");
                        return;
                    }
                    var nextWaypoint = NormalizePlanText(action.PlanWaypoint, 160);
                    var revised = requestedStatus == "revised";
                    var hasStructuredIdentity =
                        plannedInputs.Length > 0 || nextInputs.Length > 0;
                    var changed = hasStructuredIdentity
                        ? !plannedInputs.SequenceEqual(
                              nextInputs,
                              StringComparer.Ordinal) ||
                          !string.Equals(
                              waypoint,
                              nextWaypoint,
                              StringComparison.Ordinal)
                        : !string.Equals(plan, nextPlan, StringComparison.Ordinal);
                    var preserveProgress = HasActivePlan && !revised && !changed;
                    plan = nextPlan;
                    status = "active";
                    revisionReason = revised ? nextReason : null;
                    planConfidence = action.PlanConfidence ?? action.Confidence ?? planConfidence;
                    if (!preserveProgress)
                    {
                        plannedInputs = nextInputs;
                        waypoint = nextWaypoint;
                        currentInputIndex = 0;
                        originStateId = currentStateId;
                        expectedStateId = currentStateId;
                        var requestedStateId = NormalizePlanText(action.PlanStateId, 32);
                        if (requestedStateId is not null &&
                            currentStateId is not null &&
                            !string.Equals(
                                requestedStateId,
                                currentStateId,
                                StringComparison.Ordinal))
                        {
                            Invalidate(
                                $"the proposed route was bound to stale state {requestedStateId}, current state is {currentStateId}");
                            return;
                        }
                    }
                    if (changed || revised)
                    {
                        Console.WriteLine(
                            revised
                                ? $"[plan] revised: {plan}; reason={revisionReason ?? "new evidence"}"
                                : $"[plan] active: {plan}");
                        if (plannedInputs.Length > 0)
                        {
                            Console.WriteLine(
                                $"[turn-plan] state={originStateId ?? "unknown"}; waypoint={waypoint ?? "unspecified"}; confidence={(planConfidence is double confidence ? confidence.ToString("0.00") : "n/a")}; inputs=[{string.Join(",", plannedInputs)}]");
                        }
                    }
                }

                internal bool ProposedSequenceMatches(
                    IReadOnlyList<string> proposedInputs,
                    string? currentStateId,
                    out string reason)
                {
                    reason = "";
                    if (!HasExecutableDirectionalPlan)
                        return true;
                    if (expectedStateId is not null &&
                        currentStateId is not null &&
                        !string.Equals(expectedStateId, currentStateId, StringComparison.Ordinal))
                    {
                        reason =
                            $"structured route expects state {expectedStateId}, current state is {currentStateId}";
                        return false;
                    }

                    var normalized = proposedInputs
                        .Select(NormalizePlannedInput)
                        .ToArray();
                    if (normalized.Any(input => input is null))
                    {
                        reason = "structured route contains a non-directional input";
                        return false;
                    }
                    var remaining = plannedInputs.Skip(currentInputIndex).ToArray();
                    if (normalized.Length > remaining.Length ||
                        !normalized.Select(input => input!).SequenceEqual(
                            remaining.Take(normalized.Length),
                            StringComparer.Ordinal))
                    {
                        reason =
                            $"proposed inputs do not match remaining structured route [{string.Join(",", remaining)}]";
                        return false;
                    }
                    return true;
                }

                internal bool TryExpandDirectionalSequence(
                    IReadOnlyList<string> proposedInputs,
                    string? currentStateId,
                    int maximumLength,
                    out string[] expandedInputs)
                {
                    expandedInputs = proposedInputs.ToArray();
                    if (!HasExecutableDirectionalPlan ||
                        EffectiveConfidence < TurnBasedTransitionTracker.MinimumStructuredPlanConfidence ||
                        maximumLength < 2 ||
                        !ProposedSequenceMatches(
                            proposedInputs,
                            currentStateId,
                            out _))
                    {
                        return false;
                    }

                    var remaining = plannedInputs
                        .Skip(currentInputIndex)
                        .Take(maximumLength)
                        .ToArray();
                    if (remaining.Length <= proposedInputs.Count)
                        return false;
                    expandedInputs = remaining;
                    return true;
                }

                internal bool IsExpectedInput(string input, string? currentStateId)
                {
                    if (!HasExecutableDirectionalPlan)
                        return false;
                    var normalized = NormalizePlannedInput(input);
                    return normalized is not null &&
                           string.Equals(
                               TurnInput.Direction(plannedInputs[currentInputIndex]),
                               TurnInput.Direction(normalized),
                               StringComparison.Ordinal) &&
                           (expectedStateId is null || currentStateId is null ||
                            string.Equals(
                                expectedStateId,
                                currentStateId,
                                StringComparison.Ordinal));
                }

                internal void RecordDirectionalResult(
                    string input,
                    string beforeState,
                    string afterState,
                    bool changed)
                {
                    if (!HasExecutableDirectionalPlan)
                        return;
                    if (!IsExpectedInput(input, beforeState))
                    {
                        Invalidate("the executed directional input diverged from the structured route");
                        return;
                    }
                    if (!changed)
                    {
                        Invalidate("a structured route input was blocked or had no effect");
                        return;
                    }

                    currentInputIndex++;
                    expectedStateId = afterState;
                    if (currentInputIndex >= plannedInputs.Length)
                    {
                        status = "completed";
                        revisionReason = "the concrete route prefix reached its waypoint";
                        Console.WriteLine(
                            $"[turn-plan] completed {currentInputIndex}/{plannedInputs.Length} input(s); waypoint={waypoint ?? "unspecified"}.");
                        return;
                    }
                    Console.WriteLine(
                        $"[turn-plan] advanced to {currentInputIndex}/{plannedInputs.Length}; remaining=[{string.Join(",", plannedInputs.Skip(currentInputIndex))}]");
                }

                internal void Invalidate(string reason)
                {
                    if (!HasActivePlan)
                        return;
                    status = "invalidated";
                    revisionReason = NormalizePlanText(reason, 200);
                    Console.WriteLine(
                        $"[plan] invalidated: {revisionReason ?? "observed state contradicted the plan"}");
                }

                internal string BuildPromptSummary()
                {
                    if (plan is null || status == "none")
                        return "";
                    var builder = new StringBuilder()
                        .AppendLine($"SHORT_TERM_PLAN_STATUS: {status}")
                        .AppendLine($"SHORT_TERM_PLAN: {plan}");
                    if (plannedInputs.Length > 0)
                    {
                        builder
                            .AppendLine($"PLAN_STATE_ID: {originStateId ?? "unknown"}")
                            .AppendLine($"PLAN_EXPECTED_STATE_ID: {expectedStateId ?? "unknown"}")
                            .AppendLine($"PLAN_WAYPOINT: {waypoint ?? "unspecified"}")
                            .AppendLine($"PLAN_CONFIDENCE: {(planConfidence is double confidence ? confidence.ToString("0.00") : "N/A")}")
                            .AppendLine($"CURRENT_PLAN_INDEX: {currentInputIndex}/{plannedInputs.Length}")
                            .AppendLine($"PLANNED_INPUTS: [{string.Join(",", plannedInputs)}]")
                            .AppendLine($"REMAINING_PLANNED_INPUTS: [{string.Join(",", plannedInputs.Skip(currentInputIndex))}]");
                    }
                    if (!string.IsNullOrWhiteSpace(revisionReason))
                        builder.Append($"SHORT_TERM_PLAN_REVISION_REASON: {revisionReason}");
                    return builder.ToString();
                }

                static string[] NormalizePlannedInputs(string[]? values) =>
                    values?
                        .Select(NormalizePlannedInput)
                        .Where(value => value is not null)
                        .Select(value => value!)
                        .Take(Options.TurnBasedMaxBatchInputs)
                        .ToArray() ?? [];

                static string? NormalizePlannedInput(string? value) => TurnInput.Key(value);

                static string? NormalizePlanText(string? value, int maxChars)
                {
                    if (string.IsNullOrWhiteSpace(value))
                        return null;
                    return TrimForMeta(value.Trim(), maxChars);
                }
            }
