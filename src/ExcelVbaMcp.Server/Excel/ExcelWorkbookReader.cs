using System.Runtime.InteropServices;
namespace ExcelVbaMcp.Excel;

/// <summary>Enumerates metadata from attached Excel applications without changing them.</summary>
internal interface IExcelWorkbookReader
{
    ListWorkbooksResponse ListWorkbooks();
}

internal sealed class ExcelWorkbookReader(IExcelInstanceLocator instanceLocator) : IExcelWorkbookReader
{
    public ListWorkbooksResponse ListWorkbooks()
    {
        IReadOnlyList<IExcelApplication> applications;
        try
        {
            applications = instanceLocator.FindRunningApplications();
        }
        catch (COMException exception)
        {
            throw new ExcelComException("Excel is running but its workbooks could not be read.", exception);
        }

        if (applications.Count == 0)
        {
            return new ListWorkbooksResponse(false, []);
        }

        List<WorkbookInfo> workbooks = [];
        try
        {
            foreach (IExcelApplication application in applications)
            {
                application.ReadWorkbooks(workbooks);
            }

            return new ListWorkbooksResponse(true, workbooks);
        }
        catch (COMException exception)
        {
            throw new ExcelComException("Excel is running but its workbooks could not be read.", exception);
        }
        finally
        {
            for (int index = applications.Count - 1; index >= 0; index--)
            {
                applications[index].Dispose();
            }
        }
    }

}
