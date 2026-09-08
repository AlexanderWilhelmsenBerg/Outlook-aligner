using System.Globalization;
using System.Runtime.InteropServices;
using OutlookAligner.Outlook.Contracts;
using OutlookAligner.OutlookHost.Com;
using OutlookInterop = Microsoft.Office.Interop.Outlook;

namespace OutlookAligner.OutlookHost.Calendars;

internal static class ManagedCopyMetadataReader
{
    internal static ManagedCopyMetadataDto Read(
        OutlookInterop.AppointmentItem appointment,
        ICollection<string> warnings)
    {
        ArgumentNullException.ThrowIfNull(appointment);
        ArgumentNullException.ThrowIfNull(warnings);

        OutlookInterop.UserProperties? properties = null;

        try
        {
            properties = appointment.UserProperties;
            if (properties is null)
            {
                return Empty(ManagedCopyState.None);
            }

            return Classify(
                ReadString(properties, OutlookAlignerManagedProperties.SyncGroupId),
                ReadString(properties, OutlookAlignerManagedProperties.SourceGlobalAppointmentId),
                ReadString(properties, OutlookAlignerManagedProperties.SourceAccountId),
                ReadString(properties, OutlookAlignerManagedProperties.CopyType),
                ReadString(properties, OutlookAlignerManagedProperties.SchemaVersion));
        }
        catch (COMException exception)
        {
            warnings.Add($"Outlook Aligner managed-copy metadata could not be read for one item (HRESULT 0x{exception.ErrorCode:X8}).");
            return Empty(ManagedCopyState.Unreadable);
        }
        finally
        {
            ComRelease.Release(properties);
        }
    }

    internal static ManagedCopyMetadataDto Classify(
        string? syncGroupId,
        string? sourceGlobalAppointmentId,
        string? sourceAccountId,
        string? copyType,
        string? schemaVersion)
    {
        syncGroupId = Normalize(syncGroupId);
        sourceGlobalAppointmentId = Normalize(sourceGlobalAppointmentId);
        sourceAccountId = Normalize(sourceAccountId);
        copyType = Normalize(copyType);
        schemaVersion = Normalize(schemaVersion);

        if (syncGroupId is null
            && sourceGlobalAppointmentId is null
            && sourceAccountId is null
            && copyType is null
            && schemaVersion is null)
        {
            return Empty(ManagedCopyState.None);
        }

        if (schemaVersion is not null
            && !string.Equals(
                schemaVersion,
                OutlookAlignerManagedProperties.CurrentSchemaVersion,
                StringComparison.Ordinal))
        {
            return new ManagedCopyMetadataDto(
                ManagedCopyState.UnsupportedSchema,
                syncGroupId,
                sourceGlobalAppointmentId,
                sourceAccountId,
                copyType,
                schemaVersion);
        }

        var validCopyType = string.Equals(copyType, OutlookAlignerManagedProperties.CopyTypeFull, StringComparison.Ordinal)
                            || string.Equals(copyType, OutlookAlignerManagedProperties.CopyTypeBusy, StringComparison.Ordinal);

        var complete = Guid.TryParse(syncGroupId, out _)
                       && sourceGlobalAppointmentId is not null
                       && sourceAccountId is not null
                       && validCopyType
                       && string.Equals(
                           schemaVersion,
                           OutlookAlignerManagedProperties.CurrentSchemaVersion,
                           StringComparison.Ordinal);

        return new ManagedCopyMetadataDto(
            complete ? ManagedCopyState.Valid : ManagedCopyState.Incomplete,
            syncGroupId,
            sourceGlobalAppointmentId,
            sourceAccountId,
            copyType,
            schemaVersion);
    }

    private static string? ReadString(OutlookInterop.UserProperties properties, string name)
    {
        OutlookInterop.UserProperty? property = null;

        try
        {
            property = properties.Find(name, Custom: true);
            if (property is null)
            {
                return null;
            }

            return Normalize(Convert.ToString(property.Value, CultureInfo.InvariantCulture));
        }
        finally
        {
            ComRelease.Release(property);
        }
    }

    private static string? Normalize(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static ManagedCopyMetadataDto Empty(ManagedCopyState state)
        => new(state, null, null, null, null, null);
}
