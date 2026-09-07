using System.Runtime.InteropServices;
using OutlookAligner.Outlook.Contracts;
using OutlookAligner.OutlookHost.Calendars;
using OutlookAligner.OutlookHost.Com;
using OutlookInterop = Microsoft.Office.Interop.Outlook;

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

        OutlookInterop.Application? application = null;
        OutlookInterop.NameSpace? session = null;
        OutlookInterop.Stores? outlookStores = null;
        OutlookInterop.Accounts? outlookAccounts = null;

        try
        {
            application = new OutlookInterop.Application();
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
                OutlookInterop.Store? store = null;
                try
                {
                    store = outlookStores[index];
                    if (store is null)
                    {
                        warnings.Add("One Outlook store resolved to null and was skipped.");
                        continue;
                    }

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
                OutlookInterop.Account? account = null;
                OutlookInterop.Store? deliveryStore = null;
                OutlookInterop.MAPIFolder? calendar = null;

                try
                {
                    account = outlookAccounts[index];
                    if (account is null)
                    {
                        warnings.Add("One Outlook account resolved to null and was skipped.");
                        continue;
                    }

                    var liveAccount = (OutlookInterop.Account)account;
                    var displayName = liveAccount.DisplayName;
                    var smtpAddress = ReadSmtpAddress(liveAccount, warnings);
                    var accountType = liveAccount.AccountType.ToString();
                    string? storeDisplayName = null;
                    string? storeId = null;

                    try
                    {
                        deliveryStore = liveAccount.DeliveryStore;
                        if (deliveryStore is null)
                        {
                            accounts.Add(UnavailableAccount(
                                displayName,
                                smtpAddress,
                                accountType,
                                storeDisplayName,
                                storeId,
                                "Delivery store unavailable."));
                            continue;
                        }

                        storeDisplayName = deliveryStore.DisplayName;
                        storeId = deliveryStore.StoreID;
                        calendar = deliveryStore.GetDefaultFolder(OutlookInterop.OlDefaultFolders.olFolderCalendar);
                        if (calendar is null)
                        {
                            accounts.Add(UnavailableAccount(
                                displayName,
                                smtpAddress,
                                accountType,
                                storeDisplayName,
                                storeId,
                                "Default calendar unavailable."));
                            continue;
                        }

                        var liveCalendar = (OutlookInterop.MAPIFolder)calendar;
                        var events = OutlookCalendarReader.Read(
                            liveCalendar,
                            windowStartLocal,
                            windowEndLocal,
                            options.IncludeDetails,
                            warnings);

                        accounts.Add(new OutlookAccountDto(
                            displayName,
                            smtpAddress,
                            accountType,
                            storeDisplayName,
                            storeId,
                            liveCalendar.EntryID,
                            CalendarAvailable: true,
                            events.Count,
                            events,
                            Error: null));
                    }
                    catch (COMException exception)
                    {
                        accounts.Add(UnavailableAccount(
                            displayName,
                            smtpAddress,
                            accountType,
                            storeDisplayName,
                            storeId,
                            $"Calendar access/read failed (HRESULT 0x{exception.ErrorCode:X8})."));
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

    private static OutlookAccountDto UnavailableAccount(
        string displayName,
        string? smtpAddress,
        string accountType,
        string? storeDisplayName,
        string? storeId,
        string error)
        => new(
            displayName,
            smtpAddress,
            accountType,
            storeDisplayName,
            storeId,
            CalendarEntryId: null,
            CalendarAvailable: false,
            EventCount: 0,
            Events: [],
            Error: error);

    private static string? ReadSmtpAddress(OutlookInterop.Account account, ICollection<string> warnings)
    {
        try
        {
            var smtpAddress = account.SmtpAddress;
            return string.IsNullOrWhiteSpace(smtpAddress) ? null : smtpAddress;
        }
        catch (COMException exception)
        {
            warnings.Add($"One account SMTP address could not be read (HRESULT 0x{exception.ErrorCode:X8}).");
            return null;
        }
    }
}
