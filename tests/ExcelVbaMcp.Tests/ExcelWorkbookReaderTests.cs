using System.Runtime.InteropServices;
using ExcelVbaMcp.Excel;

namespace ExcelVbaMcp.Tests;

public sealed class ExcelWorkbookReaderTests
{
    [Fact]
    public async Task RunningObjectTableLocator_DoesNotRequireExcelToBeOpen()
    {
        using var dispatcher = new ExcelComDispatcher();
        IReadOnlyList<IExcelApplication> applications = await dispatcher.InvokeAsync(
            new ExcelInstanceLocator().FindRunningApplications,
            CancellationToken.None);
        try
        {
            Assert.NotNull(applications);
        }
        finally
        {
            for (int index = applications.Count - 1; index >= 0; index--)
            {
                applications[index].Dispose();
            }
        }
    }

    [Fact]
    public void NoRunningInstance_ReturnsExcelNotRunningAndAnEmptyCollection()
    {
        var reader = new ExcelWorkbookReader(new FakeLocator([]));

        ListWorkbooksResponse response = reader.ListWorkbooks();

        Assert.False(response.ExcelRunning);
        Assert.Empty(response.Workbooks);
    }

    [Fact]
    public void EmptyRunningInstance_ReturnsExcelRunningAndAnEmptyCollection()
    {
        var application = new FakeApplication();
        var reader = new ExcelWorkbookReader(new FakeLocator([application]));

        ListWorkbooksResponse response = reader.ListWorkbooks();

        Assert.True(response.ExcelRunning);
        Assert.Empty(response.Workbooks);
        Assert.True(application.Disposed);
    }

    [Fact]
    public void NeverSavedWorkbook_ReportsANullPathAndSavedFalse()
    {
        // Excel can report Workbook.Saved=true for a newly created untouched BookN even
        // though it has no path. The COM adapter normalizes that case before it reaches
        // the reader; this test locks down the resulting public contract.
        var application = new FakeApplication(new WorkbookInfo("Book3", null, false, false));
        var reader = new ExcelWorkbookReader(new FakeLocator([application]));

        ListWorkbooksResponse response = reader.ListWorkbooks();

        WorkbookInfo workbook = Assert.Single(response.Workbooks);
        Assert.Equal("Book3", workbook.Name);
        Assert.Null(workbook.FullPath);
        Assert.False(workbook.Saved);
        Assert.True(application.Disposed);
    }

    [Fact]
    public void PopulatedInstances_PreserveUnsavedPathsAndDuplicateNames()
    {
        var first = new FakeApplication(
            new WorkbookInfo("Budget.xlsx", @"C:\Finance\Budget.xlsx", true, false),
            new WorkbookInfo("Book1", null, false, false));
        var second = new FakeApplication(
            new WorkbookInfo("Budget.xlsx", @"D:\Archive\Budget.xlsx", true, true));
        var reader = new ExcelWorkbookReader(new FakeLocator([first, second]));

        ListWorkbooksResponse response = reader.ListWorkbooks();

        Assert.True(response.ExcelRunning);
        Assert.Collection(
            response.Workbooks,
            workbook => Assert.Equal(new WorkbookInfo("Budget.xlsx", @"C:\Finance\Budget.xlsx", true, false), workbook),
            workbook => Assert.Equal(new WorkbookInfo("Book1", null, false, false), workbook),
            workbook => Assert.Equal(new WorkbookInfo("Budget.xlsx", @"D:\Archive\Budget.xlsx", true, true), workbook));
        Assert.True(first.Disposed);
        Assert.True(second.Disposed);
    }

    [Fact]
    public void LocatorComFailure_IsTranslatedToUsefulExcelComError()
    {
        var reader = new ExcelWorkbookReader(new FakeLocator(new COMException("Access denied")));

        ExcelComException exception = Assert.Throws<ExcelComException>(reader.ListWorkbooks);

        Assert.IsType<COMException>(exception.InnerException);
        Assert.Contains("workbooks could not be read", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EnumerationFailure_DisposesEveryAcquiredApplicationInReverseOrder()
    {
        var events = new List<string>();
        var first = new FakeApplication(events, "first", new WorkbookInfo("One.xlsx", @"C:\One.xlsx", true, false));
        var second = new FakeApplication(events, "second", new COMException("Workbook collection unavailable"));
        var reader = new ExcelWorkbookReader(new FakeLocator([first, second]));

        ExcelComException exception = Assert.Throws<ExcelComException>(reader.ListWorkbooks);

        Assert.IsType<COMException>(exception.InnerException);
        Assert.Equal(["read:first", "read:second", "dispose:second", "dispose:first"], events);
        Assert.True(first.Disposed);
        Assert.True(second.Disposed);
    }

    private sealed class FakeLocator : IExcelInstanceLocator
    {
        private readonly IReadOnlyList<IExcelApplication>? applications;
        private readonly Exception? exception;

        public FakeLocator(IReadOnlyList<IExcelApplication> applications) => this.applications = applications;

        public FakeLocator(Exception exception) => this.exception = exception;

        public IReadOnlyList<IExcelApplication> FindRunningApplications()
        {
            if (exception is not null)
            {
                throw exception;
            }

            return applications!;
        }
    }

    private sealed class FakeApplication : IExcelApplication
    {
        private readonly IReadOnlyList<WorkbookInfo>? workbooks;
        private readonly Exception? readException;
        private readonly List<string>? events;
        private readonly string? id;

        public FakeApplication(params WorkbookInfo[] workbooks) => this.workbooks = workbooks;

        public FakeApplication(List<string> events, string id, params WorkbookInfo[] workbooks)
        {
            this.events = events;
            this.id = id;
            this.workbooks = workbooks;
        }

        public FakeApplication(List<string> events, string id, Exception readException)
        {
            this.events = events;
            this.id = id;
            this.readException = readException;
        }

        public bool Disposed { get; private set; }

        public void ReadWorkbooks(ICollection<WorkbookInfo> results)
        {
            events?.Add($"read:{id}");
            if (readException is not null)
            {
                throw readException;
            }

            foreach (WorkbookInfo workbook in workbooks!)
            {
                results.Add(workbook);
            }
        }

        public void Dispose()
        {
            Disposed = true;
            events?.Add($"dispose:{id}");
        }
    }
}
