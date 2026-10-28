namespace RDPilot.Control;

/// <summary>Owns bounded visual prototypes and their associated navigation data.</summary>
internal sealed class TurnStateStore
{
    internal const int StateLimit = 512;
    internal const int EdgeLimit = 128;
    internal sealed class Prototype(string id, byte[] fingerprint)
    {
        internal string Id { get; } = id;
        internal byte[] Fingerprint { get; set; } = fingerprint;
    }
    internal readonly record struct NavigationEdge(string FromState, Point From, string Action,
        string ToState, Point To, bool Blocked);
    readonly List<Prototype> states = [];
    readonly Dictionary<string, Point> positions = new(StringComparer.Ordinal);
    readonly Queue<NavigationEdge> edges = new();
    internal IReadOnlyList<Prototype> States => states;
    internal IReadOnlyDictionary<string, Point> Positions => positions;
    internal IReadOnlyCollection<NavigationEdge> Edges => edges;

    internal void Clear() { states.Clear(); ClearNavigation(); }
    internal void ClearNavigation() { positions.Clear(); edges.Clear(); }
    internal void Anchor(string id)
    {
        if (states.Any(state => state.Id == id)) positions.TryAdd(id, Point.Empty);
    }
    internal void Add(Prototype state, string? currentState)
    {
        states.Add(state);
        while (states.Count > StateLimit)
        {
            var removed = states.First(item => item.Id != currentState && item.Id != state.Id);
            states.Remove(removed);
            positions.Remove(removed.Id);
            var retainedEdges = edges.Where(edge => edge.FromState != removed.Id && edge.ToState != removed.Id).ToArray();
            edges.Clear();
            foreach (var edge in retainedEdges) edges.Enqueue(edge);
        }
    }
    internal void RecordTransition(string before, string action, string after, string result, string? currentState)
    {
        var delta = action switch
        {
            "ArrowUp" or "W" => new Point(0, -1), "ArrowDown" or "S" => new Point(0, 1),
            "ArrowLeft" or "A" => new Point(-1, 0), "ArrowRight" or "D" => new Point(1, 0),
            _ => Point.Empty
        };
        if (delta == Point.Empty || !states.Any(state => state.Id == before) ||
            !states.Any(state => state.Id == after)) return;
        if (!positions.TryGetValue(before, out var from))
        {
            from = positions.GetValueOrDefault(currentState ?? "");
            positions[before] = from;
        }
        var blocked = result == "no_effect";
        var to = blocked ? from : positions.TryGetValue(after, out var known) ? known :
            new Point(from.X + delta.X, from.Y + delta.Y);
        if (!blocked) positions[after] = to;
        edges.Enqueue(new(before, from, action, after, to, blocked));
        while (edges.Count > EdgeLimit) edges.Dequeue();
    }
}
