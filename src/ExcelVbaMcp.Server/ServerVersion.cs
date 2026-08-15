using System.Reflection;

namespace ExcelVbaMcp;

internal static class ServerVersion
{
    public static string Current { get; } = Resolve();

    private static string Resolve()
    {
        Assembly assembly = typeof(ServerVersion).Assembly;
        string? informationalVersion = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion;

        return informationalVersion?.Split('+', 2)[0]
            ?? assembly.GetName().Version?.ToString(3)
            ?? "unknown";
    }
}
