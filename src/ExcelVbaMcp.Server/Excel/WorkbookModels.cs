using System.Text.Json.Serialization;

namespace ExcelVbaMcp.Excel;

/// <summary>Plain workbook data returned by the server; it never retains COM objects.</summary>
public sealed record WorkbookInfo(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("fullPath")] string? FullPath,
    [property: JsonPropertyName("saved")] bool Saved,
    [property: JsonPropertyName("readOnly")] bool ReadOnly);

/// <summary>Result of enumerating already-running Excel instances.</summary>
public sealed record ListWorkbooksResponse(
    [property: JsonPropertyName("excelRunning")] bool ExcelRunning,
    [property: JsonPropertyName("workbooks")] IReadOnlyList<WorkbookInfo> Workbooks);

internal sealed class ExcelComException : Exception
{
    public ExcelComException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}

internal sealed class ExcelOperationTimeoutException : TimeoutException
{
    public ExcelOperationTimeoutException(TimeSpan timeout)
        : base($"Excel did not respond within {timeout.TotalSeconds:0} seconds. It may be busy or displaying a modal dialog.")
    {
    }
}
