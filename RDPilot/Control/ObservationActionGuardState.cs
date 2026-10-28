
namespace RDPilot.Control;


            internal sealed class ObservationActionGuardState
            {
                readonly HashSet<string> inspectionSignatures = new(StringComparer.Ordinal);

                internal int ConsecutiveInspectionActions { get; private set; }
                internal bool AimIssuedSinceInteraction { get; private set; }
                internal Rectangle? AimRect { get; private set; }

                internal void SetAim(Rectangle region)
                {
                    AimRect = region;
                    AimIssuedSinceInteraction = true;
                }

                internal void InvalidateAim()
                {
                    AimRect = null;
                    AimIssuedSinceInteraction = false;
                }

                internal bool RequiresInteraction(int maxConsecutiveInspectionActions) =>
                    maxConsecutiveInspectionActions > 0 &&
                    ConsecutiveInspectionActions >= maxConsecutiveInspectionActions;

                internal bool TryGetBlockReason(
                    ResolvedActionSnapshot action,
                    int maxConsecutiveInspectionActions,
                    out string reason) =>
                    TryGetBlockReason(
                        action,
                        maxConsecutiveInspectionActions,
                        singleInspectionBeforeInteraction: false,
                        out reason);

                internal bool TryGetBlockReason(
                    ResolvedActionSnapshot action,
                    int maxConsecutiveInspectionActions,
                    bool singleInspectionBeforeInteraction,
                    out string reason)
                {
                    reason = "";
                    if (IsInspectionAction(action.Action))
                    {
                        if (AimIssuedSinceInteraction)
                        {
                            reason = "an AIM is already active; interact with that target instead of requesting another inspection";
                            return true;
                        }

                        var signature = InspectionSignature(action.Action);
                        if (inspectionSignatures.Contains(signature))
                        {
                            reason = "the requested inspection revisits an area already inspected since the last interaction; choose a state-changing action";
                            return true;
                        }

                        if (singleInspectionBeforeInteraction &&
                            ConsecutiveInspectionActions >= 1)
                        {
                            reason = "a detailed turn-based inspection was already supplied; perform a state-changing input before requesting another nested crop";
                            return true;
                        }

                        if (RequiresInteraction(maxConsecutiveInspectionActions))
                        {
                            reason = $"the limit of {maxConsecutiveInspectionActions} consecutive crop/point inspections was reached; perform one safe state-changing interaction before inspecting again";
                            return true;
                        }
                    }
                    else if (action.Action.Type == "aim" && AimIssuedSinceInteraction)
                    {
                        reason = "an AIM is already active; click or perform the intended gesture instead of aiming again";
                        return true;
                    }

                    return false;
                }

                internal void RecordExecuted(ResolvedActionSnapshot action)
                {
                    if (IsInspectionAction(action.Action))
                    {
                        inspectionSignatures.Add(InspectionSignature(action.Action));
                        ConsecutiveInspectionActions++;
                        return;
                    }

                    // AIM is registered with its resolved rectangle by SetAim during execution.
                    if (IsStateChangingInteractionAction(action.Action))
                    {
                        ResetInspectionBudget();
                        // Keep the geometric click context, but allow a new AIM after interaction.
                        AimIssuedSinceInteraction = false;
                    }
                }

                internal void Reset()
                {
                    ResetInspectionBudget();
                    InvalidateAim();
                }

                void ResetInspectionBudget()
                {
                    inspectionSignatures.Clear();
                    ConsecutiveInspectionActions = 0;
                }

                static string InspectionSignature(ActionDto action)
                {
                    var rect = ResolveCropRect(action);
                    if (rect is not Rectangle region)
                        return IneffectiveActionSignature(action);

                    const int quantum = 16;
                    return $"{action.Type}:{region.Left / quantum},{region.Top / quantum},{region.Width / quantum},{region.Height / quantum}";
                }
            }
