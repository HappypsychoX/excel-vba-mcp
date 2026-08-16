using System.ComponentModel;
using ExcelVbaMcp.Excel;
using ModelContextProtocol.Server;

namespace ExcelVbaMcp.Tools;

[McpServerToolType]
internal sealed class WorkbookTools(ExcelComDispatcher dispatcher, IExcelWorkbookReader workbookReader)
{
    private static readonly TimeSpan OperationTimeout = TimeSpan.FromSeconds(10);

    [McpServerTool(Name = "list_workbooks", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("Lists workbooks already open in Excel without creating, opening, saving, closing, or quitting Excel.")]
    public async Task<ListWorkbooksResponse> ListWorkbooksAsync(CancellationToken cancellationToken)
    {
        using CancellationTokenSource timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(OperationTimeout);

        try
        {
            return await dispatcher.InvokeAsync(workbookReader.ListWorkbooks, timeout.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            throw new ExcelOperationTimeoutException(OperationTimeout);
        }
    }
}
