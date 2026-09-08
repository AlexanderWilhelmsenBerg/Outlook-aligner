using System.Reflection;
using OutlookInterop = Microsoft.Office.Interop.Outlook;

namespace OutlookAligner.OutlookHost.Com;

internal static class InteropDependencyCheck
{
    internal static string Run()
    {
        ValidateType(typeof(OutlookInterop._Application));
        ValidateType(typeof(OutlookInterop._NameSpace));
        ValidateType(typeof(OutlookInterop.MAPIFolder));
        ValidateType(typeof(OutlookInterop.AppointmentItem));

        return "Outlook interop metadata resolved successfully.";
    }

    private static void ValidateType(Type type)
    {
        foreach (var property in type.GetProperties(BindingFlags.Instance | BindingFlags.Public))
        {
            _ = property.PropertyType.FullName;
        }

        foreach (var method in type.GetMethods(BindingFlags.Instance | BindingFlags.Public))
        {
            _ = method.ReturnType.FullName;
            foreach (var parameter in method.GetParameters())
            {
                _ = parameter.ParameterType.FullName;
            }
        }
    }
}
