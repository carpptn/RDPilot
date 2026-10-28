namespace RDPilot.Control;

internal static class ControlLoopGuards
{
    internal static (int RepeatCount, string? LastSignature) UpdateRepeatDetection(
        ResolvedActionSnapshot previousAction,
        bool noChange,
        int repeatCount,
        string? lastSignature,
        List<ResolvedActionSnapshot> recentIneffectiveSpatialActions)
    {
        var family = RecoveryMemoryService.ActionFamily(previousAction.Action);
        var isSpatial = IsSpatialPointerAction(previousAction.Action) || family == "drag_drop";
        if (previousAction.ScreenPoint is Point pointerPoint && isSpatial)
        {
            if (noChange)
            {
                var repeatsNearbyPoint = recentIneffectiveSpatialActions.Any(prior =>
                {
                    if (!string.Equals(RecoveryMemoryService.ActionFamily(prior.Action), family, StringComparison.OrdinalIgnoreCase) ||
                        MouseButton(prior.Action) != MouseButton(previousAction.Action) ||
                        prior.ScreenPoint is not Point priorPoint ||
                        !ScreenPointsAreNearby(priorPoint, pointerPoint))
                    {
                        return false;
                    }

                    if (family != "drag_drop")
                        return true;
                    return prior.DestinationScreenPoint is Point priorDestination &&
                        previousAction.DestinationScreenPoint is Point destination &&
                        ScreenPointsAreNearby(priorDestination, destination);
                });
                repeatCount = repeatsNearbyPoint ? repeatCount + 1 : 0;
                recentIneffectiveSpatialActions.Add(previousAction);
                if (recentIneffectiveSpatialActions.Count > 16)
                    recentIneffectiveSpatialActions.RemoveAt(0);
            }
            else
            {
                repeatCount = 0;
                recentIneffectiveSpatialActions.Clear();
            }

            return (repeatCount, null);
        }

        recentIneffectiveSpatialActions.Clear();
        if (!noChange)
            return (0, null);
        var signature = previousAction.IneffectiveSignature;
        return (string.Equals(signature, lastSignature, StringComparison.Ordinal) ? repeatCount + 1 : 0, signature);
    }

    internal static int RepeatedStringCycleLength(IReadOnlyList<string> values)
    {
        for (var period = 1; period <= Math.Min(8, values.Count / 2); period++)
        {
            var start = values.Count - period * 2;
            var matches = true;
            for (var index = 0; index < period; index++)
            {
                if (string.Equals(values[start + index], values[start + period + index], StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                matches = false;
                break;
            }

            if (matches)
                return period;
        }

        return 0;
    }

    internal static bool ScreenPointsAreNearby(Point a, Point b)
    {
        var radius = Math.Max(16, Options.IneffectiveMouseClusterPx);
        var dx = (long)a.X - b.X;
        var dy = (long)a.Y - b.Y;
        return dx * dx + dy * dy <= (long)radius * radius;
    }

    internal static bool IsActionOnCooldown(
        ResolvedActionSnapshot action,
        int step,
        Dictionary<string, int> cooldowns,
        List<SpatialActionCooldown> spatialCooldowns,
        out int untilStep)
    {
        untilStep = 0;
        if (Options.ActionRepeatCooldownSteps <= 0 || IsLocalObservationAction(action.Action))
            return false;
        for (var i = spatialCooldowns.Count - 1; i >= 0; i--)
        {
            if (spatialCooldowns[i].UntilStep < step)
                spatialCooldowns.RemoveAt(i);
        }

        foreach (var expired in cooldowns.Where(item => item.Value < step).Select(item => item.Key).ToArray())
        {
            cooldowns.Remove(expired);
        }

        if (cooldowns.TryGetValue(action.IneffectiveSignature, out untilStep))
        {
            if (step <= untilStep)
                return true;
            cooldowns.Remove(action.IneffectiveSignature);
        }

        var family = RecoveryMemoryService.ActionFamily(action.Action);
        if (action.ScreenPoint is Point screenPoint &&
            (IsSpatialPointerAction(action.Action) || family == "drag_drop"))
        {
            foreach (var cooldown in spatialCooldowns)
            {
                if (!string.Equals(cooldown.ActionFamily, family, StringComparison.OrdinalIgnoreCase) ||
                    cooldown.Button != MouseButton(action.Action))
                    continue;
                if (!ScreenPointsAreNearby(cooldown.ScreenPoint, screenPoint))
                    continue;
                if (family == "drag_drop" &&
                    (cooldown.DestinationScreenPoint is not Point priorDestination ||
                    action.DestinationScreenPoint is not Point destination ||
                    !ScreenPointsAreNearby(priorDestination, destination)))
                {
                    continue;
                }

                untilStep = Math.Max(untilStep, cooldown.UntilStep);
            }

            return untilStep >= step;
        }

        return false;
    }

    internal static void RegisterActionCooldown(
        ResolvedActionSnapshot action,
        int untilStep,
        Dictionary<string, int> cooldowns,
        List<SpatialActionCooldown> spatialCooldowns,
        bool clusterSpatially = true)
    {
        var family = RecoveryMemoryService.ActionFamily(action.Action);
        if (clusterSpatially &&
            action.ScreenPoint is Point screenPoint &&
            (IsSpatialPointerAction(action.Action) ||
            family == "drag_drop"))
        {
            spatialCooldowns.Add(new SpatialActionCooldown(screenPoint, action.DestinationScreenPoint, family, untilStep, MouseButton(action.Action)));
            if (spatialCooldowns.Count > Math.Max(16, Options.RuntimeCooldownEntryLimit))
            {
                spatialCooldowns.RemoveRange(0, spatialCooldowns.Count - Math.Max(16, Options.RuntimeCooldownEntryLimit));
            }

            return;
        }

        cooldowns[action.IneffectiveSignature] = untilStep;
        if (cooldowns.Count > Math.Max(16, Options.RuntimeCooldownEntryLimit))
        {
            foreach (var key in cooldowns.OrderBy(item => item.Value).Take(cooldowns.Count - Math.Max(16, Options.RuntimeCooldownEntryLimit)).Select(item => item.Key)
                .ToArray())
            {
                cooldowns.Remove(key);
            }
        }
    }
}
