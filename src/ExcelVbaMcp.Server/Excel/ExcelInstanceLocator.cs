using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using OfficeExcel = Microsoft.Office.Interop.Excel;

namespace ExcelVbaMcp.Excel;

/// <summary>Finds Excel applications already registered in the Windows Running Object Table.</summary>
internal interface IExcelInstanceLocator
{
    IReadOnlyList<IExcelApplication> FindRunningApplications();
}

internal interface IExcelApplication : IDisposable
{
    void ReadWorkbooks(ICollection<WorkbookInfo> results);
}

internal sealed class ExcelInstanceLocator : IExcelInstanceLocator
{
    public IReadOnlyList<IExcelApplication> FindRunningApplications()
    {
        ComTypes.IRunningObjectTable? runningObjectTable = null;
        ComTypes.IBindCtx? bindContext = null;
        ComTypes.IEnumMoniker? monikerEnumerator = null;
        List<IExcelApplication> applications = [];
        HashSet<int> instanceHandles = [];

        try
        {
            int result = GetRunningObjectTable(0, out runningObjectTable);
            Marshal.ThrowExceptionForHR(result);
            result = CreateBindCtx(0, out bindContext);
            Marshal.ThrowExceptionForHR(result);
            runningObjectTable.EnumRunning(out monikerEnumerator);

            ComTypes.IMoniker[] monikers = new ComTypes.IMoniker[1];
            while (monikerEnumerator.Next(1, monikers, IntPtr.Zero) == 0)
            {
                ComTypes.IMoniker moniker = monikers[0];
                object? candidate = null;
                OfficeExcel._Application? application = null;
                bool applicationTransferred = false;
                try
                {
                    try
                    {
                        runningObjectTable.GetObject(moniker, out candidate);
                    }
                    catch (COMException)
                    {
                        // A ROT can contain objects owned by unrelated applications. Their
                        // failure must not be reported as an inaccessible Excel instance.
                        continue;
                    }

                    application = GetApplication(candidate);
                    if (application is null)
                    {
                        continue;
                    }

                    int handle = application.Hwnd;
                    if (!instanceHandles.Add(handle))
                    {
                        continue;
                    }

                    applications.Add(new ExcelApplicationReference(application, handle.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                    applicationTransferred = true;
                }
                catch (COMException exception)
                {
                    throw new ExcelComException("Excel is running but its instance could not be accessed.", exception);
                }
                finally
                {
                    // A workbook ROT entry yields a separate application reference. The original
                    // entry must always be released before advancing to the next moniker.
                    if (!ReferenceEquals(candidate, application))
                    {
                        ComReleaser.Release(candidate);
                    }

                    if (!applicationTransferred)
                    {
                        ComReleaser.Release(application);
                    }
                    ComReleaser.Release(moniker);
                }
            }

            return applications;
        }
        catch
        {
            ReleaseApplications(applications);
            throw;
        }
        finally
        {
            ComReleaser.Release(monikerEnumerator);
            ComReleaser.Release(bindContext);
            ComReleaser.Release(runningObjectTable);
        }
    }

    private static OfficeExcel._Application? GetApplication(object candidate)
    {
        if (candidate is OfficeExcel._Application application)
        {
            return application;
        }

        if (candidate is OfficeExcel._Workbook workbook)
        {
            return workbook.Application;
        }

        return null;
    }

    private static void ReleaseApplications(List<IExcelApplication> applications)
    {
        for (int index = applications.Count - 1; index >= 0; index--)
        {
            applications[index].Dispose();
        }
    }

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(uint reserved, out ComTypes.IRunningObjectTable runningObjectTable);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out ComTypes.IBindCtx bindContext);
}

/// <summary>Owns one ROT-acquired Excel application reference until the caller disposes it.</summary>
internal sealed class ExcelApplicationReference(OfficeExcel._Application application, string instanceId) : IExcelApplication
{
    public OfficeExcel._Application Application { get; } = application;

    public string InstanceId { get; } = instanceId;

    public void ReadWorkbooks(ICollection<WorkbookInfo> results)
    {
        OfficeExcel.Workbooks? workbookCollection = null;
        try
        {
            workbookCollection = Application.Workbooks;
            int workbookCount = workbookCollection.Count;
            for (int index = 1; index <= workbookCount; index++)
            {
                OfficeExcel._Workbook? workbook = null;
                try
                {
                    workbook = workbookCollection[index];
                    string path = workbook.Path;
                    bool hasSavedPath = !string.IsNullOrWhiteSpace(path);
                    results.Add(new WorkbookInfo(
                        workbook.Name,
                        hasSavedPath ? workbook.FullName : null,
                        hasSavedPath && workbook.Saved,
                        workbook.ReadOnly));
                }
                finally
                {
                    ComReleaser.Release(workbook);
                }
            }
        }
        finally
        {
            ComReleaser.Release(workbookCollection);
        }
    }

    public void Dispose() => ComReleaser.Release(Application);
}
