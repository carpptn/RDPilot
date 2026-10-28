using System.Reflection;

namespace RDPilot.SelfTests.Support;

internal static class TestState
{
    internal sealed class Slot(FieldInfo metadata, object? target)
    {
        internal string Name => metadata.Name;
        internal Type FieldType => metadata.FieldType;
        internal bool IsLiteral => metadata.IsLiteral;
        internal bool IsInitOnly => metadata.IsInitOnly;
        internal object? GetValue(object? ignored) => metadata.GetValue(target);
        internal void SetValue(object? ignored, object? value) => metadata.SetValue(target, value);
    }

    internal static Slot Field(Type type, string name, BindingFlags flags)
    {
        if (type == typeof(RDPilotApplication))
        {
            foreach (var owner in new object[] { RDPilotApplication.Options, RDPilotApplication.CurrentRun })
            {
                var metadata = owner.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
                if (metadata is not null) return new(metadata, owner);
            }
        }
        return new(type.GetField(name, flags) ?? throw new MissingFieldException(type.Name, name), null);
    }

    internal static IEnumerable<Slot> Fields(Type type, BindingFlags flags) =>
        type == typeof(RDPilotApplication)
            ? typeof(AppOptions).GetFields(BindingFlags.Instance | BindingFlags.NonPublic)
                .Select(metadata => new Slot(metadata, RDPilotApplication.Options))
            : type.GetFields(flags).Select(metadata => new Slot(metadata, null));
}
