using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;

if (args.Length != 1 || !File.Exists(args[0]))
{
    Console.Error.WriteLine("Usage: dotnet run --project tests/ExcelVbaMcp.SmokeTest -- <server-executable>");
    return 2;
}

string serverPath = Path.GetFullPath(args[0]);
var transport = new StdioClientTransport(new StdioClientTransportOptions
{
    Name = "ExcelVbaMcp Phase 1 smoke test",
    Command = serverPath,
    ShutdownTimeout = TimeSpan.FromSeconds(5),
});

await using McpClient client = await McpClient.CreateAsync(transport);

IList<McpClientTool> tools = await client.ListToolsAsync();
string[] actualNames = [.. tools.Select(tool => tool.Name).Order()];
string[] expectedNames = ["get_version", "ping"];

if (!actualNames.SequenceEqual(expectedNames))
{
    throw new InvalidOperationException(
        $"Expected exactly [{string.Join(", ", expectedNames)}], got [{string.Join(", ", actualNames)}].");
}

string ping = await ReadTextResultAsync(client, "ping");
if (ping != "pong")
{
    throw new InvalidOperationException($"Expected ping to return 'pong', got '{ping}'.");
}

string version = await ReadTextResultAsync(client, "get_version");
if (!Version.TryParse(version, out _))
{
    throw new InvalidOperationException($"Expected a parseable server version, got '{version}'.");
}

Console.WriteLine($"Listed exactly: {string.Join(", ", actualNames)}");
Console.WriteLine($"ping => {ping}");
Console.WriteLine($"get_version => {version}");
return 0;

static async Task<string> ReadTextResultAsync(McpClient client, string toolName)
{
    CallToolResult result = await client.CallToolAsync(
        toolName,
        new Dictionary<string, object?>(),
        cancellationToken: CancellationToken.None);

    return result.Content.OfType<TextContentBlock>().Single().Text;
}
