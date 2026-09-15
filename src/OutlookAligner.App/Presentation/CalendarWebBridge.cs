using System.Text.Json;
using System.Text.Json.Serialization;
using OutlookAligner.App.ViewModels;

namespace OutlookAligner.App.Presentation;

public static class CalendarWebProtocol
{
    public const int Version = 1;

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public static string SerializeRender(IReadOnlyList<CalendarPresentationObservation> observations)
        => JsonSerializer.Serialize(new CalendarRenderMessage(Version, "renderObservations", observations), JsonOptions);

    public static string SerializeNavigation(string action)
        => JsonSerializer.Serialize(new CalendarNavigationMessage(Version, "navigate", action), JsonOptions);

    public static bool TryParseClientMessage(string json, out CalendarClientMessage? message)
    {
        message = null;
        try
        {
            using var document = JsonDocument.Parse(json);
            var root = document.RootElement;
            if (!root.TryGetProperty("version", out var versionElement)
                || versionElement.ValueKind != JsonValueKind.Number
                || versionElement.GetInt32() != Version
                || !root.TryGetProperty("type", out var typeElement)
                || typeElement.ValueKind != JsonValueKind.String)
            {
                return false;
            }

            var type = typeElement.GetString();
            switch (type)
            {
                case "ready":
                    message = new CalendarClientMessage(CalendarClientMessageKind.Ready);
                    return true;
                case "rangeChanged":
                    if (!TryReadRequiredString(root, "title", out var title)
                        || !TryReadRequiredString(root, "start", out var start)
                        || !TryReadRequiredString(root, "end", out var end))
                    {
                        return false;
                    }

                    if (!root.TryGetProperty("visibleCount", out var visibleCountElement)
                        || visibleCountElement.ValueKind != JsonValueKind.Number
                        || !visibleCountElement.TryGetInt32(out var visibleCount)
                        || visibleCount < 0)
                    {
                        return false;
                    }

                    message = new CalendarClientMessage(
                        CalendarClientMessageKind.RangeChanged,
                        Title: title,
                        Start: start,
                        End: end,
                        VisibleCount: visibleCount);
                    return true;
                case "observationSelected":
                    if (!TryReadRequiredString(root, "presentationId", out var presentationId))
                    {
                        return false;
                    }

                    message = new CalendarClientMessage(CalendarClientMessageKind.ObservationSelected, PresentationId: presentationId);
                    return true;
                default:
                    return false;
            }
        }
        catch (JsonException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
    }

    private static bool TryReadRequiredString(JsonElement root, string name, out string value)
    {
        value = string.Empty;
        if (!root.TryGetProperty(name, out var element) || element.ValueKind != JsonValueKind.String)
        {
            return false;
        }

        value = element.GetString() ?? string.Empty;
        return !string.IsNullOrWhiteSpace(value);
    }
}

public sealed class CalendarPresentationSession
{
    private readonly Dictionary<string, CalendarEventRowViewModel> _observations = new(StringComparer.Ordinal);

    public IReadOnlyList<CalendarPresentationObservation> Reset(IEnumerable<CalendarEventRowViewModel> source)
    {
        _observations.Clear();
        var result = new List<CalendarPresentationObservation>();
        foreach (var row in source)
        {
            var presentationId = Guid.NewGuid().ToString("N");
            _observations[presentationId] = row;
            result.Add(new CalendarPresentationObservation(
                presentationId,
                row.Subject,
                row.CalendarEvent.StartLocal,
                row.CalendarEvent.EndLocal,
                row.CalendarEvent.IsAllDay,
                row.AccountName,
                row.Location,
                row.CalendarEvent.IsRecurring));
        }

        return result;
    }

    public bool TryResolve(string presentationId, out CalendarEventRowViewModel? observation)
        => _observations.TryGetValue(presentationId, out observation);
}

public sealed record CalendarPresentationObservation(
    string PresentationId,
    string Title,
    DateTime Start,
    DateTime End,
    bool AllDay,
    string AccountLabel,
    string Location,
    bool IsRecurring);

public sealed record CalendarRenderMessage(
    int Version,
    string Type,
    IReadOnlyList<CalendarPresentationObservation> Observations);

public sealed record CalendarNavigationMessage(int Version, string Type, string Action);

public enum CalendarClientMessageKind
{
    Ready,
    RangeChanged,
    ObservationSelected,
}

public sealed record CalendarClientMessage(
    CalendarClientMessageKind Kind,
    string? PresentationId = null,
    string? Title = null,
    string? Start = null,
    string? End = null,
    int? VisibleCount = null);
