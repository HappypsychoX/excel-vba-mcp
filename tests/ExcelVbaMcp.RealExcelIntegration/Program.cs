using System.Diagnostics;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;

IntegrationArguments options = IntegrationArguments.Parse(args);
HashSet<int> initialExcelProcessIds = GetExcelProcessIds();

Console.WriteLine($"Excel processes before setup: {FormatProcessIds(initialExcelProcessIds)}");
Console.WriteLine("Open Excel yourself, then open each saved workbook and create the requested unsaved workbook.");
Console.WriteLine("This harness never starts, closes, saves, or otherwise modifies Excel or a workbook.");
WaitForEnter("Press Enter only after Excel is ready.");

HashSet<int> excelProcessIdsBeforeServer = GetExcelProcessIds();
if (excelProcessIdsBeforeServer.Count == 0)
{
    throw new InvalidOperationException("Excel must be running before the integration call.");
}

var transport = new StdioClientTransport(new StdioClientTransportOptions
{
    Name = "ExcelVbaMcp real Excel integration harness",
    Command = options.ServerPath,
    ShutdownTimeout = TimeSpan.FromSeconds(5),
});

await using (McpClient client = await McpClient.CreateAsync(transport))
{
    ListWorkbooksResponse first = await ListWorkbooksAsync(client);
    ListWorkbooksResponse second = await ListWorkbooksAsync(client);

    AssertExpectedWorkbooks(first, options);
    AssertExpectedWorkbooks(second, options);
}

HashSet<int> excelProcessIdsAfterSession = GetExcelProcessIds();
if (!excelProcessIdsAfterSession.SetEquals(excelProcessIdsBeforeServer))
{
    throw new InvalidOperationException(
        $"Excel process IDs changed while the MCP session ran. Before: {FormatProcessIds(excelProcessIdsBeforeServer)}; " +
        $"after: {FormatProcessIds(excelProcessIdsAfterSession)}.");
}

Console.WriteLine("Two list_workbooks calls succeeded and the Excel process set was unchanged.");
Console.WriteLine("Verify the workbooks are still open and usable, then close Excel manually.");
WaitForEnter("Press Enter after manually closing the Excel instances used for this test.");

HashSet<int> remainingExcelProcessIds = GetExcelProcessIds();
HashSet<int> processesOpenedForTest = [.. excelProcessIdsBeforeServer.Except(initialExcelProcessIds)];
if (processesOpenedForTest.Overlaps(remainingExcelProcessIds))
{
    throw new InvalidOperationException(
        $"Excel process(es) opened for this test are still running: {FormatProcessIds(processesOpenedForTest.Intersect(remainingExcelProcessIds))}.");
}

Console.WriteLine("Real Excel lifecycle integration test passed.");

static async Task<ListWorkbooksResponse> ListWorkbooksAsync(McpClient client)
{
    CallToolResult result = await client.CallToolAsync(
        "list_workbooks",
        new Dictionary<string, object?>(),
        cancellationToken: CancellationToken.None);

    string json = result.Content.OfType<TextContentBlock>().Single().Text;
    return JsonSerializer.Deserialize<ListWorkbooksResponse>(json, JsonSerializerOptions.Web)
        ?? throw new InvalidOperationException("list_workbooks returned an empty JSON response.");
}

static void AssertExpectedWorkbooks(ListWorkbooksResponse response, IntegrationArguments options)
{
    if (!response.ExcelRunning)
    {
        throw new InvalidOperationException("list_workbooks reported that Excel is not running.");
    }

    foreach (string expectedName in options.SavedWorkbookNames)
    {
        WorkbookResponse workbook = FindSingle(response.Workbooks, expectedName);
        if (workbook.FullPath is null || !workbook.Saved)
        {
            throw new InvalidOperationException($"Saved workbook '{expectedName}' did not include a saved full path.");
        }
    }

    WorkbookResponse unsavedWorkbook = FindSingle(response.Workbooks, options.UnsavedWorkbookName);
    if (unsavedWorkbook.FullPath is not null || unsavedWorkbook.Saved)
    {
        throw new InvalidOperationException(
            $"Unsaved workbook '{options.UnsavedWorkbookName}' must have fullPath null and saved false.");
    }
}

static WorkbookResponse FindSingle(IEnumerable<WorkbookResponse> workbooks, string name)
{
    WorkbookResponse[] matches = [.. workbooks.Where(workbook => string.Equals(workbook.Name, name, StringComparison.Ordinal))];
    return matches.Length switch
    {
        1 => matches[0],
        0 => throw new InvalidOperationException($"Workbook '{name}' was not returned by list_workbooks."),
        _ => throw new InvalidOperationException($"Workbook '{name}' was returned more than once."),
    };
}

static HashSet<int> GetExcelProcessIds() =>
    [.. Process.GetProcessesByName("EXCEL").Select(process => process.Id)];

static string FormatProcessIds(IEnumerable<int> processIds) =>
    string.Join(", ", processIds.Order()) is { Length: > 0 } formatted ? formatted : "(none)";

static void WaitForEnter(string prompt)
{
    Console.WriteLine(prompt);
    _ = Console.ReadLine();
}

internal sealed record ListWorkbooksResponse(bool ExcelRunning, List<WorkbookResponse> Workbooks);

internal sealed record WorkbookResponse(string Name, string? FullPath, bool Saved, bool ReadOnly);

internal sealed record IntegrationArguments(string ServerPath, IReadOnlyList<string> SavedWorkbookNames, string UnsavedWorkbookName)
{
    public static IntegrationArguments Parse(string[] args)
    {
        string? serverPath = null;
        List<string> savedWorkbookNames = [];
        string? unsavedWorkbookName = null;

        for (int index = 0; index < args.Length; index++)
        {
            if (index + 1 >= args.Length)
            {
                throw new ArgumentException($"Missing value for '{args[index]}'.");
            }

            string value = args[++index];
            switch (args[index - 1])
            {
                case "--server":
                    serverPath = value;
                    break;
                case "--workbook":
                    savedWorkbookNames.Add(value);
                    break;
                case "--unsaved":
                    unsavedWorkbookName = value;
                    break;
                default:
                    throw new ArgumentException($"Unknown option '{args[index - 1]}'.");
            }
        }

        if (serverPath is null || !File.Exists(serverPath) || savedWorkbookNames.Count != 2 || string.IsNullOrWhiteSpace(unsavedWorkbookName))
        {
            throw new ArgumentException(
                "Usage: dotnet run --project tests/ExcelVbaMcp.RealExcelIntegration -- " +
                "--server <ExcelVbaMcp.exe> --workbook <saved-name-1> --workbook <saved-name-2> --unsaved <unsaved-name>");
        }

        return new IntegrationArguments(Path.GetFullPath(serverPath), savedWorkbookNames, unsavedWorkbookName);
    }
}
