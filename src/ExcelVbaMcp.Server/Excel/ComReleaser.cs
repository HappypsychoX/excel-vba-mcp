using System.Runtime.InteropServices;

namespace ExcelVbaMcp.Excel;

/// <summary>Releases COM references acquired by one dispatcher operation.</summary>
internal static class ComReleaser
{
    public static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            _ = Marshal.ReleaseComObject(value);
        }
    }
}
