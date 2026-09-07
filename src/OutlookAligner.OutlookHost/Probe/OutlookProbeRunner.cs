using System.Runtime.InteropServices;
using OutlookAligner.Outlook.Contracts;
using OutlookAligner.OutlookHost.Calendars;
using OutlookAligner.OutlookHost.Com;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAligner.OutlookHost.Probe;

internal static class OutlookProbeRunner
{
    internal static OutlookProbeResult Run(ProbeOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        var windowStartLocal = DateTime.Today;
        var windowEndLocal = windowStartLocal.AddDays(options.Days);
        var warnings = new List<string>();
        var stores = new List<OutlookStoreDto>();
        var accounts = new List<OutlookAccountDto>();

        Outlook.Application? application = null;
        Outlook.NameSpace? session = null;
        Outlook.Stores? outlookStores = null;
        Outlook.Accounts? outlookAccounts = null;

        try
        {
            application = new Outlook.Application();
            session = application.GetNamespace("MAPI");

            string? profileName;
            try
            {
                profileName = session.CurrentProfileName;
            }
            catch (COMException exception)
            {
                profileName = null;
                warnings.Add($"Profile name could not be read (HRESULT 0x{exception.ErrorCode:X8}).");
            }

            outlookStores = session.Stores;
            for (var index = 1; index <= outlookStores.Count; index++)
            {
                Outlook.Store? store = null;
                try
                {
                    store = outlookStores[index];
                    stores.Add(new OutlookStoreDto(store.DisplayName, store.StoreID));
                }
                catch (COMException exception)
                {
                    warnings.Add($"One Outlook store could not be inspected (HRESULT 0x{exception.ErrorCode:X8}).");
                }
                finally
                {
                    ComRelease.Release(store);
                }
            }

            outlookAccounts = session.Accounts;
            for (var index = 1; index <= outlookAccounts.Count; index++)
            {
                Outlook.Account? account = null;
                Outlook.Store? deliveryStore = null;
                Outlook.MAPIFolder? calendar = null;

                try
                {
                    account = outlookAccounts[index];
                    var displayName = account.DisplayName;
                    var smtpAddress = ReadSmtpAddress(account, warnings);
                    var accountType = account.AccountType.ToString();

                    try
                    {
                        deliveryStore = account.DeliveryStore;
                        calendar = deliveryStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                        var events = OutlookCalendarReader.Read(
                            calendar,
                            windowStartLocal,
                            windowEndLocal,
                            options.IncludeDetails,
                            warnings);

                        accounts.Add(new OutlookAccountDto(
                            displayName,
                            smtpAddress,
                            accountType,
                            deliveryStore.DisplayName,
                            deliveryStore.StoreID,
                            calendar.EntryID,
                            CalendarAvailable: true,
                            events.Count,
                            events,
                            Error: null));
                    }
                    catch (COMException exception)
                    {
                        var error = $"Default calendar unavailable (HRESULT 0x{exception.ErrorCode:X8}).";
                        accounts.Add(new OutlookAccountDto(
                            displayName,
                            smtpAddress,
                            accountType,
                            deliveryStore?.DisplayName,
                            deliveryStore?.StoreID,
                            CalendarEntryId: null,
                            CalendarAvailable: false,
                            EventCount: 0,
                            Events: [],
                            Error: error));
                    }
                }
                catch (COMException exception)
                {
                    warnings.Add($"One Outlook account could not be inspected (HRESULT 0x{exception.ErrorCode:X8}).");
                }
                finally
                {
                    ComRelease.Release(calendar);
                    ComRelease.Release(deliveryStore);
                    ComRelease.Release(account);
                }
            }

            return new OutlookProbeResult(
                OutlookHostProtocol.Version,
                DateTime.UtcNow,
                windowStartLocal,
                windowEndLocal,
                profileName,
                stores,
                accounts,
                warnings);
        }
        finally
        {
            ComRelease.Release(outlookAccounts);
            ComRelease.Release(outlookStores);
            ComRelease.Release(session);
            ComRelease.Release(application);
        }
    }

    private static string? ReadSmtpAddress(Outlook.Account account, ICollection<string> warnings)
    {
        try
        {
            return string.IsNullOrWhiteSpace(account.SmtpAddress) ? null : account.SmtpAddress;
        }
        catch (COMException exception)
        {
            warnings.Add($"One account SMTP address could not be read (HRESULT 0x{exception.ErrorCode:X8}).");
            return null;
        }
    }
}
