using System.Globalization;
using System.Runtime.InteropServices;
using OutlookAligner.Outlook.Contracts;
using OutlookAligner.OutlookHost.Com;
using OutlookInterop = Microsoft.Office.Interop.Outlook;

namespace OutlookAligner.OutlookHost.Calendars;

internal static class OutlookCalendarReader
{
    internal static IReadOnlyList<CalendarEventDto> Read(
        OutlookInterop.MAPIFolder calendar,
        DateTime windowStartLocal,
        DateTime windowEndLocal,
        bool includeDetails,
        ICollection<string> warnings)
    {
        ArgumentNullException.ThrowIfNull(calendar);
        ArgumentNullException.ThrowIfNull(warnings);

        OutlookInterop.Items? items = null;
        OutlookInterop.Items? restrictedItems = null;
        object? current = null;
        var events = new List<CalendarEventDto>();

        try
        {
            items = calendar.Items;
            items.Sort("[Start]");
            items.IncludeRecurrences = true;

            var restriction = OutlookDateFilter.Build(
                windowStartLocal,
                windowEndLocal,
                CultureInfo.CurrentCulture);
            restrictedItems = items.Restrict(restriction);

            current = restrictedItems.GetFirst();
            while (current is not null)
            {
                object? next = null;
                try
                {
                    if (current is OutlookInterop.AppointmentItem appointment)
                    {
                        try
                        {
                            events.Add(ToDto(appointment, includeDetails, warnings));
                        }
                        catch (COMException exception)
                        {
                            warnings.Add(
                                $"One calendar item could not be read (HRESULT 0x{exception.ErrorCode:X8}).");
                        }
                    }
                }
                finally
                {
                    next = restrictedItems.GetNext();
                    ComRelease.Release(current);
                    current = next;
                }
            }
        }
        finally
        {
            ComRelease.Release(current);
            ComRelease.Release(restrictedItems);
            ComRelease.Release(items);
        }

        return events;
    }

    private static CalendarEventDto ToDto(
        OutlookInterop.AppointmentItem appointment,
        bool includeDetails,
        ICollection<string> warnings)
        => new(
            appointment.EntryID,
            appointment.GlobalAppointmentID,
            appointment.Start,
            appointment.End,
            appointment.AllDayEvent,
            appointment.IsRecurring,
            appointment.RecurrenceState.ToString(),
            appointment.MeetingStatus.ToString(),
            appointment.BusyStatus.ToString(),
            appointment.Sensitivity.ToString(),
            includeDetails ? appointment.Subject : null,
            includeDetails ? appointment.Location : null,
            ManagedCopyMetadataReader.Read(appointment, warnings));
}
