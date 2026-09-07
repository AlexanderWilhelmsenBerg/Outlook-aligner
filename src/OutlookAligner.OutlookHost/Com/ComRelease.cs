using System.Runtime.InteropServices;

namespace OutlookAligner.OutlookHost.Com;

internal static class ComRelease
{
    internal static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value))
        {
            return;
        }

        _ = Marshal.ReleaseComObject(value);
    }
}
