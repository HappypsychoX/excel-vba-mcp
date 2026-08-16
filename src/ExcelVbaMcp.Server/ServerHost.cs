using ExcelVbaMcp.Tools;
using ExcelVbaMcp.Excel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;

namespace ExcelVbaMcp;

internal static class ServerHost
{
    public static async Task RunAsync(string[] args)
    {
        HostApplicationBuilder builder = Host.CreateApplicationBuilder(args);

        // stdout is reserved for MCP protocol messages when using stdio.
        builder.Logging.AddConsole(options =>
            options.LogToStandardErrorThreshold = LogLevel.Trace);

        builder.Services
            .AddMcpServer()
            .WithStdioServerTransport()
            .WithTools<PhaseOneTools>()
            .WithTools<WorkbookTools>();

        builder.Services.AddSingleton<ExcelComDispatcher>();
        builder.Services.AddSingleton<IExcelInstanceLocator, ExcelInstanceLocator>();
        builder.Services.AddSingleton<IExcelWorkbookReader, ExcelWorkbookReader>();

        await builder.Build().RunAsync();
    }
}
