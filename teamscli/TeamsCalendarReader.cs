using System.Net.Http.Headers;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace teamscli;

internal sealed class TeamsCalendarReader
{
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";
    private const int EventsPageSize = 50;
    private const int MaxConsecutiveEmptyEventPages = 5;
    private static readonly TimeSpan DefaultCalendarRangeBeforeNow = TimeSpan.FromDays(180);
    private static readonly TimeSpan DefaultCalendarRangeAfterNow = TimeSpan.FromDays(365);

    public async Task ReadAsync(
        string token,
        int top,
        bool includeBody,
        string subjectFilter,
        string organizerFilter,
        string startAfter,
        string endBefore)
    {
        var events = await ListEventsAsync(token, top, includeBody, subjectFilter, organizerFilter, startAfter, endBefore);
        PrintEvents(events, includeBody);
    }

    public async Task ReadEventAsync(string token, string eventId, bool includeBody)
    {
        var calendarEvent = await GetEventAsync(token, eventId, includeBody, null, null);
        PrintEvent(calendarEvent, includeBody);
    }

    public async Task ReadEventAsync(string token, string eventId, bool includeBody, string startAfter, string endBefore)
    {
        var calendarEvent = await GetEventAsync(token, eventId, includeBody, startAfter, endBefore);
        PrintEvent(calendarEvent, includeBody);
    }

    private async Task<List<JsonElement>> ListEventsAsync(
        string token,
        int top,
        bool includeBody,
        string subjectFilter,
        string organizerFilter,
        string startAfter,
        string endBefore)
    {
        var selectFields = includeBody
            ? "id,subject,start,end,organizer,attendees,location,body,webLink,type,seriesMasterId"
            : "id,subject,start,end,organizer,attendees,location,webLink,type,seriesMasterId";

        var (calendarViewStart, calendarViewEnd) = GetCalendarViewRange(startAfter, endBefore);

        var query = new Dictionary<string, string>
        {
            ["$top"] = EventsPageSize.ToString(),
            ["$select"] = selectFields,
            ["$orderby"] = "start/dateTime desc",
            ["startDateTime"] = calendarViewStart,
            ["endDateTime"] = calendarViewEnd
        };

        var url = BuildUrl($"{GraphBaseUrl}/me/calendar/calendarView", query);

        using var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (includeBody)
        {
            client.DefaultRequestHeaders.Add("Prefer", "outlook.body-content-type=\"text\"");
        }

        var items = new List<JsonElement>();
        var consecutiveEmptyPages = 0;

        while (!string.IsNullOrWhiteSpace(url) && items.Count < top)
        {
            using var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {content}");
            }

            using var doc = JsonDocument.Parse(content);
            var root = doc.RootElement;
            var pageHasData = false;

            if (root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in valueElement.EnumerateArray())
                {
                    pageHasData = true;

                    if (!MatchesEventFilters(item, subjectFilter, organizerFilter))
                    {
                        continue;
                    }

                    items.Add(item.Clone());
                    if (items.Count >= top)
                    {
                        break;
                    }
                }
            }

            if (pageHasData)
            {
                consecutiveEmptyPages = 0;
            }
            else
            {
                consecutiveEmptyPages++;
                if (consecutiveEmptyPages >= MaxConsecutiveEmptyEventPages)
                {
                    break;
                }
            }

            url = root.TryGetProperty("@odata.nextLink", out var nextElement) ? nextElement.GetString() : null;
        }

        return items;
    }

    private static (string StartDateTime, string EndDateTime) GetCalendarViewRange(string startAfter, string endBefore)
    {
        var utcNow = DateTimeOffset.UtcNow;
        var hasStart = TryParseUtcDateTime(startAfter, out var parsedStart);
        var hasEnd = TryParseUtcDateTime(endBefore, out var parsedEnd);

        if (hasStart && hasEnd)
        {
            return (ToGraphUtcDateTimeString(parsedStart), ToGraphUtcDateTimeString(parsedEnd));
        }

        if (hasStart)
        {
            return (ToGraphUtcDateTimeString(parsedStart), ToGraphUtcDateTimeString(parsedStart.Add(DefaultCalendarRangeAfterNow)));
        }

        if (hasEnd)
        {
            return (ToGraphUtcDateTimeString(parsedEnd.Subtract(DefaultCalendarRangeBeforeNow)), ToGraphUtcDateTimeString(parsedEnd));
        }

        return (
            ToGraphUtcDateTimeString(utcNow.Subtract(DefaultCalendarRangeBeforeNow)),
            ToGraphUtcDateTimeString(utcNow.Add(DefaultCalendarRangeAfterNow)));
    }

    private static bool MatchesEventFilters(JsonElement ev, string subjectFilter, string organizerFilter)
    {
        if (!string.IsNullOrWhiteSpace(subjectFilter))
        {
            var subject = GetString(ev, "subject") ?? string.Empty;
            if (subject.IndexOf(subjectFilter.Trim(), StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }
        }

        if (!string.IsNullOrWhiteSpace(organizerFilter))
        {
            var normalizedFilter = organizerFilter.Trim();
            var organizerName = string.Empty;
            var organizerAddress = string.Empty;

            if (ev.TryGetProperty("organizer", out var organizerObj)
                && organizerObj.ValueKind == JsonValueKind.Object
                && organizerObj.TryGetProperty("emailAddress", out var emailAddressObj)
                && emailAddressObj.ValueKind == JsonValueKind.Object)
            {
                organizerName = GetString(emailAddressObj, "name") ?? string.Empty;
                organizerAddress = GetString(emailAddressObj, "address") ?? string.Empty;
            }

            var nameMatch = organizerName.IndexOf(normalizedFilter, StringComparison.OrdinalIgnoreCase) >= 0;
            var addressMatch = organizerAddress.IndexOf(normalizedFilter, StringComparison.OrdinalIgnoreCase) >= 0;
            if (!nameMatch && !addressMatch)
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryParseUtcDateTime(string value, out DateTimeOffset parsed)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            parsed = default;
            return false;
        }

        if (!DateTimeOffset.TryParse(value, out parsed))
        {
            parsed = default;
            return false;
        }

        parsed = parsed.ToUniversalTime();
        return true;
    }

    private static string ToGraphUtcDateTimeString(DateTimeOffset value)
    {
        return value.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
    }

    private async Task<JsonElement> GetEventAsync(string token, string eventId, bool includeBody, string startAfter, string endBefore)
    {
        var selectFields = includeBody
            ? "id,subject,start,end,organizer,attendees,location,body,webLink,type,seriesMasterId,onlineMeeting"
            : "id,subject,start,end,organizer,attendees,location,webLink,type,seriesMasterId,onlineMeeting";

        var url = BuildUrl(
            $"{GraphBaseUrl}/me/events/{Uri.EscapeDataString(eventId)}",
            new Dictionary<string, string>
            {
                ["$select"] = selectFields
            });

        using var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (includeBody)
        {
            client.DefaultRequestHeaders.Add("Prefer", "outlook.body-content-type=\"text\"");
        }

        using var response = await client.GetAsync(url);
        var content = await response.Content.ReadAsStringAsync();
        if (response.IsSuccessStatusCode)
        {
            using var doc = JsonDocument.Parse(content);
            var calendarEvent = doc.RootElement.Clone();
            var eventType = GetString(calendarEvent, "type");
            if (string.Equals(eventType, "seriesMaster", StringComparison.OrdinalIgnoreCase))
            {
                var recurringInstance = await ResolveEventFromSeriesInstancesAsync(client, eventId, includeBody, startAfter, endBefore);
                if (recurringInstance.HasValue)
                {
                    return recurringInstance.Value;
                }
            }

            return calendarEvent;
        }

        if (response.StatusCode is not System.Net.HttpStatusCode.NotFound and not System.Net.HttpStatusCode.BadRequest)
        {
            throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {content}");
        }

        var fallback = await ResolveEventFromCalendarViewAsync(client, eventId, includeBody, startAfter, endBefore);
        if (fallback.HasValue)
        {
            return fallback.Value;
        }

        throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {content}");
    }

    private async Task<JsonElement?> ResolveEventFromSeriesInstancesAsync(HttpClient client, string seriesMasterEventId, bool includeBody, string startAfter, string endBefore)
    {
        if (string.IsNullOrWhiteSpace(startAfter) && string.IsNullOrWhiteSpace(endBefore))
        {
            return null;
        }

        var selectFields = includeBody
            ? "id,subject,start,end,organizer,attendees,location,body,webLink,type,seriesMasterId,onlineMeeting"
            : "id,subject,start,end,organizer,attendees,location,webLink,type,seriesMasterId,onlineMeeting";

        var (rangeStart, rangeEnd) = GetCalendarViewRange(startAfter, endBefore);
        var query = new Dictionary<string, string>
        {
            ["$top"] = EventsPageSize.ToString(),
            ["$select"] = selectFields,
            ["$orderby"] = "start/dateTime asc",
            ["startDateTime"] = rangeStart,
            ["endDateTime"] = rangeEnd
        };

        var url = BuildUrl($"{GraphBaseUrl}/me/events/{Uri.EscapeDataString(seriesMasterEventId)}/instances", query);
        JsonElement? firstInstance = null;
        var count = 0;

        while (!string.IsNullOrWhiteSpace(url))
        {
            using var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {content}");
            }

            using var doc = JsonDocument.Parse(content);
            var root = doc.RootElement;
            if (root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in valueElement.EnumerateArray())
                {
                    count++;
                    firstInstance ??= item.Clone();
                    if (count > 1)
                    {
                        break;
                    }
                }
            }

            if (count > 1)
            {
                break;
            }

            url = root.TryGetProperty("@odata.nextLink", out var nextElement) ? nextElement.GetString() : null;
        }

        if (count == 1 && firstInstance.HasValue)
        {
            return firstInstance;
        }

        if (count > 1)
        {
            throw new InvalidOperationException("Multiple recurring instances matched this event id. Specify a narrower --start and --end window.");
        }

        return null;
    }

    private async Task<JsonElement?> ResolveEventFromCalendarViewAsync(HttpClient client, string eventId, bool includeBody, string startAfter, string endBefore)
    {
        var selectFields = includeBody
            ? "id,subject,start,end,organizer,attendees,location,body,webLink,type,seriesMasterId,onlineMeeting"
            : "id,subject,start,end,organizer,attendees,location,webLink,type,seriesMasterId,onlineMeeting";

        var (calendarViewStart, calendarViewEnd) = GetCalendarViewRange(startAfter, endBefore);
        var query = new Dictionary<string, string>
        {
            ["$top"] = EventsPageSize.ToString(),
            ["$select"] = selectFields,
            ["$orderby"] = "start/dateTime asc",
            ["startDateTime"] = calendarViewStart,
            ["endDateTime"] = calendarViewEnd
        };

        var url = BuildUrl($"{GraphBaseUrl}/me/calendar/calendarView", query);
        JsonElement? firstSeriesOccurrence = null;
        var seriesOccurrenceCount = 0;

        while (!string.IsNullOrWhiteSpace(url))
        {
            using var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {content}");
            }

            using var doc = JsonDocument.Parse(content);
            var root = doc.RootElement;

            if (root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in valueElement.EnumerateArray())
                {
                    var currentId = GetString(item, "id");
                    if (string.Equals(currentId, eventId, StringComparison.Ordinal))
                    {
                        return item.Clone();
                    }

                    var seriesMasterId = GetString(item, "seriesMasterId");
                    if (string.Equals(seriesMasterId, eventId, StringComparison.Ordinal))
                    {
                        seriesOccurrenceCount++;
                        firstSeriesOccurrence ??= item.Clone();
                    }
                }
            }

            url = root.TryGetProperty("@odata.nextLink", out var nextElement) ? nextElement.GetString() : null;
        }

        if (seriesOccurrenceCount == 1 && firstSeriesOccurrence.HasValue)
        {
            return firstSeriesOccurrence;
        }

        if (seriesOccurrenceCount > 1)
        {
            throw new InvalidOperationException("Multiple recurring instances matched this event id. Specify --start and --end to narrow the occurrence window.");
        }

        return null;
    }

    private static void PrintEvents(List<JsonElement> events, bool includeBody)
    {
        var payload = new JsonArray(events.Select((ev, index) => (JsonNode)ToCalendarEvent(ev, index + 1, includeBody)).ToArray());
        Console.WriteLine(payload.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
    }

    private static void PrintEvent(JsonElement calendarEvent, bool includeBody)
    {
        var payload = ToCalendarEvent(calendarEvent, null, includeBody);
        Console.WriteLine(payload.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
    }

    private static JsonObject ToCalendarEvent(JsonElement ev, int? index, bool includeBody)
    {
        var subject = GetString(ev, "subject") ?? "(no subject)";
        var organizer = ev.TryGetProperty("organizer", out var organizerObj) && organizerObj.ValueKind == JsonValueKind.Object
            ? FormatRecipient(organizerObj)
            : null;
        var attendees = ev.TryGetProperty("attendees", out var attendeesObj) && attendeesObj.ValueKind == JsonValueKind.Array
            ? attendeesObj.EnumerateArray().Select(FormatAttendee).OfType<string>().ToList()
            : new List<string>();
        var location = ev.TryGetProperty("location", out var locationObj) && locationObj.ValueKind == JsonValueKind.Object
            ? GetString(locationObj, "displayName")
            : null;
        var startObj = ev.TryGetProperty("start", out var start) ? start : default;
        var endObj = ev.TryGetProperty("end", out var end) ? end : default;

        var attendeeArray = new JsonArray(attendees.Select(x => JsonValue.Create(x)).ToArray());
        var eventObject = new JsonObject
        {
            ["id"] = GetString(ev, "id"),
            ["subject"] = subject,
            ["organizer"] = organizer,
            ["attendees"] = attendeeArray,
            ["location"] = location,
            ["start"] = startObj.ValueKind == JsonValueKind.Object
                ? new JsonObject
                {
                    ["dateTime"] = GetString(startObj, "dateTime"),
                    ["timeZone"] = GetString(startObj, "timeZone")
                }
                : null,
            ["end"] = endObj.ValueKind == JsonValueKind.Object
                ? new JsonObject
                {
                    ["dateTime"] = GetString(endObj, "dateTime"),
                    ["timeZone"] = GetString(endObj, "timeZone")
                }
                : null,
            ["webLink"] = GetString(ev, "webLink")
        };

        if (index.HasValue)
        {
            eventObject["index"] = index.Value;
        }

        if (includeBody)
        {
            var body = ev.TryGetProperty("body", out var bodyObj) && bodyObj.ValueKind == JsonValueKind.Object
                ? new JsonObject
                {
                    ["contentType"] = GetString(bodyObj, "contentType"),
                    ["content"] = GetString(bodyObj, "content")
                }
                : null;

            eventObject["body"] = body;
        }

        return eventObject;
    }

    private static string BuildUrl(string baseUrl, Dictionary<string, string> query)
    {
        var qs = string.Join("&", query.Select(kvp =>
            $"{Uri.EscapeDataString(kvp.Key)}={Uri.EscapeDataString(kvp.Value)}"));
        return string.IsNullOrWhiteSpace(qs) ? baseUrl : $"{baseUrl}?{qs}";
    }

    private static string GetString(JsonElement element, string propertyName)
    {
        if (element.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        return element.TryGetProperty(propertyName, out var value) && value.ValueKind != JsonValueKind.Null
            ? value.GetString()
            : null;
    }

    private static string FormatRecipient(JsonElement element)
    {
        if (element.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        if (element.TryGetProperty("emailAddress", out var emailAddress) && emailAddress.ValueKind == JsonValueKind.Object)
        {
            var name = GetString(emailAddress, "name");
            var address = GetString(emailAddress, "address");

            if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(address))
            {
                return $"{name} <{address}>";
            }

            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            if (!string.IsNullOrWhiteSpace(address))
            {
                return address;
            }
        }

        return GetString(element, "displayName") ?? GetString(element, "name");
    }

    private static string FormatAttendee(JsonElement element)
    {
        return FormatRecipient(element);
    }

}
