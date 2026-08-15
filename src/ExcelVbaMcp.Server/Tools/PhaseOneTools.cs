using System.ComponentModel;
using ModelContextProtocol.Server;

namespace ExcelVbaMcp.Tools;

[McpServerToolType]
internal sealed class PhaseOneTools
{
    [McpServerTool(Name = "ping", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("Checks whether the Excel VBA MCP server is responsive.")]
    public static string Ping() => "pong";

    [McpServerTool(Name = "get_version", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("Returns the Excel VBA MCP server assembly/package version.")]
    public static string GetVersion() => ServerVersion.Current;
}
