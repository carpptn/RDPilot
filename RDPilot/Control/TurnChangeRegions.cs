namespace RDPilot.Control;

/// <summary>Pure geometry calculations for connected visual changes and peripheral regions.</summary>
internal static class TurnChangeRegions
{
    const int StateFingerprintSide = TurnVisualAnalysis.Side;
    internal const int ChangeTileSide = 4;

    internal readonly record struct ChangeRegion(
        int Left,
        int Top,
        int Right,
        int Bottom,
        int ChangedPixels);

    internal static bool RegionsFollowSamePattern(
        ChangeRegion previous,
        ChangeRegion current)
    {
        var previousWidth = Math.Max(1, previous.Right - previous.Left);
        var previousHeight = Math.Max(1, previous.Bottom - previous.Top);
        var currentWidth = Math.Max(1, current.Right - current.Left);
        var currentHeight = Math.Max(1, current.Bottom - current.Top);
        var widthSimilarity = Math.Min(previousWidth, currentWidth) /
                              (double)Math.Max(previousWidth, currentWidth);
        var heightSimilarity = Math.Min(previousHeight, currentHeight) /
                               (double)Math.Max(previousHeight, currentHeight);
        var intersectionWidth = Math.Max(
            0,
            Math.Min(previous.Right, current.Right) -
            Math.Max(previous.Left, current.Left));
        var intersectionHeight = Math.Max(
            0,
            Math.Min(previous.Bottom, current.Bottom) -
            Math.Max(previous.Top, current.Top));
        var intersection = intersectionWidth * intersectionHeight;
        var union = previousWidth * previousHeight +
                    currentWidth * currentHeight - intersection;
        var intersectionOverUnion = union <= 0
            ? 0
            : intersection / (double)union;
        var centerDistance = RegionCenterDistance(previous, current);
        var verticalOverlap = intersectionHeight /
                              (double)Math.Min(previousHeight, currentHeight);
        return intersectionOverUnion >= 0.15 ||
               widthSimilarity >= 0.50 &&
               heightSimilarity >= 0.50 &&
               centerDistance <= 0.12 ||
               verticalOverlap >= 0.60 &&
               heightSimilarity >= 0.50 &&
               centerDistance <= 0.12;
    }

    internal static double RegionCenterDistance(
        ChangeRegion left,
        ChangeRegion right)
    {
        var leftCenterX = (left.Left + left.Right) / 2.0;
        var leftCenterY = (left.Top + left.Bottom) / 2.0;
        var rightCenterX = (right.Left + right.Right) / 2.0;
        var rightCenterY = (right.Top + right.Bottom) / 2.0;
        return Math.Sqrt(
            Math.Pow(
                (leftCenterX - rightCenterX) /
                StateFingerprintSide,
                2) +
            Math.Pow(
                (leftCenterY - rightCenterY) /
                StateFingerprintSide,
                2));
    }

    internal static bool AreExpectedDirectionalMotionPair(
        string? actionLabel,
        ChangeRegion left,
        ChangeRegion right)
    {
        if (actionLabel is null ||
            !TurnInput.IsDirectionalSequence(actionLabel))
        {
            return false;
        }

        var deltaX = Math.Abs(
            (left.Left + left.Right) -
            (right.Left + right.Right)) / 2.0;
        var deltaY = Math.Abs(
            (left.Top + left.Bottom) -
            (right.Top + right.Bottom)) / 2.0;
        var vertical = actionLabel is "ArrowUp" or "ArrowDown" or "W" or "S";
        return vertical
            ? deltaY >= ChangeTileSide && deltaY > deltaX * 1.5
            : deltaX >= ChangeTileSide && deltaX > deltaY * 1.5;
    }

    internal static bool HasSpatiallySeparatedChangeRegions(
        IReadOnlyList<ChangeRegion> regions)
    {
        for (var first = 0; first < regions.Count; first++)
        for (var second = first + 1; second < regions.Count; second++)
        {
            if (RegionCenterDistance(regions[first], regions[second]) >= 0.30)
                return true;
        }
        return false;
    }

    internal static bool IsLikelyPeripheralAuxiliaryRegion(ChangeRegion region)
    {
        var width = Math.Max(1, region.Right - region.Left);
        var height = Math.Max(1, region.Bottom - region.Top);
        var outerBand = StateFingerprintSide / 8;
        var nearOuterEdge =
            region.Bottom <= outerBand ||
            region.Top >= StateFingerprintSide - outerBand ||
            region.Right <= outerBand ||
            region.Left >= StateFingerprintSide - outerBand;
        var compact =
            width * height <= StateFingerprintSide * StateFingerprintSide * 0.025 &&
            Math.Min(width, height) <= StateFingerprintSide / 8;
        return nearOuterEdge && compact;
    }

    internal static bool IsStrongPeripheralIndicatorCandidate(
        ChangeRegion region)
    {
        if (!IsLikelyPeripheralAuxiliaryRegion(region))
            return false;
        var width = Math.Max(1, region.Right - region.Left);
        var height = Math.Max(1, region.Bottom - region.Top);
        return Math.Max(width, height) >= Math.Min(width, height) * 1.75;
    }

    internal static bool IsPeripheralOnlyChange(
        IReadOnlyList<ChangeRegion> regions)
    {
        if (regions.Count == 0)
            return false;

        const int topChromeBoundary = StateFingerprintSide / 8;
        const int edgeBoundary = 2;
        return regions.All(region =>
            region.Bottom <= topChromeBoundary ||
            region.Top >= StateFingerprintSide - edgeBoundary ||
            region.Right <= edgeBoundary ||
            region.Left >= StateFingerprintSide - edgeBoundary);
    }

    internal static List<ChangeRegion> ConnectedChangeRegions(
        bool[] activeTiles,
        int tileCount)
    {
        var regions = new List<ChangeRegion>();
        var visited = new bool[activeTiles.Length];
        for (var start = 0; start < activeTiles.Length; start++)
        {
            if (!activeTiles[start] || visited[start])
                continue;
            var queue = new Queue<int>();
            queue.Enqueue(start);
            visited[start] = true;
            var minX = tileCount;
            var minY = tileCount;
            var maxX = 0;
            var maxY = 0;
            var changedTiles = 0;
            while (queue.Count > 0)
            {
                var current = queue.Dequeue();
                var x = current % tileCount;
                var y = current / tileCount;
                minX = Math.Min(minX, x);
                minY = Math.Min(minY, y);
                maxX = Math.Max(maxX, x);
                maxY = Math.Max(maxY, y);
                changedTiles++;
                for (var dy = -1; dy <= 1; dy++)
                for (var dx = -1; dx <= 1; dx++)
                {
                    if (dx == 0 && dy == 0)
                        continue;
                    var nx = x + dx;
                    var ny = y + dy;
                    if (nx < 0 || ny < 0 || nx >= tileCount || ny >= tileCount)
                        continue;
                    var neighbor = ny * tileCount + nx;
                    if (!activeTiles[neighbor] || visited[neighbor])
                        continue;
                    visited[neighbor] = true;
                    queue.Enqueue(neighbor);
                }
            }
            regions.Add(new ChangeRegion(
                minX * ChangeTileSide,
                minY * ChangeTileSide,
                (maxX + 1) * ChangeTileSide,
                (maxY + 1) * ChangeTileSide,
                changedTiles * ChangeTileSide * ChangeTileSide));
        }
        return regions
            .OrderByDescending(region => region.ChangedPixels)
            .Take(6)
            .ToList();
    }
}
