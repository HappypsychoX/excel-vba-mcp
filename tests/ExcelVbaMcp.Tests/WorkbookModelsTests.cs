using System.Text.Json;
using ExcelVbaMcp.Excel;

namespace ExcelVbaMcp.Tests;

public sealed class WorkbookModelsTests
{
    [Fact]
    public void NoRunningInstance_MapsToTheRequiredEmptyResponse()
    {
        var response = new ListWorkbooksResponse(false, []);

        Assert.False(response.ExcelRunning);
        Assert.Empty(response.Workbooks);
    }

    [Fact]
    public void EmptyRunningInstance_MapsToAnEmptyWorkbookCollection()
    {
        var response = new ListWorkbooksResponse(true, []);

        Assert.True(response.ExcelRunning);
        Assert.Empty(response.Workbooks);
    }

    [Fact]
    public void PopulatedWorkbookMetadata_PreservesDuplicateNamesAndInstanceSpecificPaths()
    {
        var response = new ListWorkbooksResponse(
            true,
            [
                new WorkbookInfo("Budget.xlsx", @"C:\Finance\Budget.xlsx", true, false),
                new WorkbookInfo("Budget.xlsx", @"D:\Archive\Budget.xlsx", true, true),
            ]);

        Assert.Collection(
            response.Workbooks,
            first =>
            {
                Assert.Equal("Budget.xlsx", first.Name);
                Assert.Equal(@"C:\Finance\Budget.xlsx", first.FullPath);
                Assert.False(first.ReadOnly);
            },
            second =>
            {
                Assert.Equal("Budget.xlsx", second.Name);
                Assert.Equal(@"D:\Archive\Budget.xlsx", second.FullPath);
                Assert.True(second.ReadOnly);
            });
    }

    [Fact]
    public void UnsavedWorkbook_SerializesANullPath()
    {
        var response = new ListWorkbooksResponse(
            true,
            [new WorkbookInfo("Book1", null, false, false)]);

        using JsonDocument json = JsonDocument.Parse(JsonSerializer.Serialize(response));
        JsonElement workbook = json.RootElement.GetProperty("workbooks")[0];

        Assert.Equal(JsonValueKind.Null, workbook.GetProperty("fullPath").ValueKind);
        Assert.False(workbook.GetProperty("saved").GetBoolean());
    }
}
