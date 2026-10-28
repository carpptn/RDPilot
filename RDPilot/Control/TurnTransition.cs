namespace RDPilot.Control;

/// <summary>An observed directed transition retained in the turn ledger.</summary>
internal readonly record struct TurnTransition(
    string From,
    string Action,
    string To,
    string Result,
    bool ReturnedToKnownState);
