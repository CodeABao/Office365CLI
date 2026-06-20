using System.Globalization;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace outlookcli;

internal sealed class MailReader
{
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";
    private const int GraphPageSize = 50;
    private static readonly IReadOnlyDictionary<string, string> WellKnownMailFolders = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["archive"] = "archive",
        ["clutter"] = "clutter",
        ["conflicts"] = "conflicts",
        ["conversationhistory"] = "conversationhistory",
        ["deleteditems"] = "deleteditems",
        ["drafts"] = "drafts",
        ["inbox"] = "inbox",
        ["junkemail"] = "junkemail",
        ["localfailures"] = "localfailures",
        ["msgfolderroot"] = "msgfolderroot",
        ["outbox"] = "outbox",
        ["recoverableitemsdeletions"] = "recoverableitemsdeletions",
        ["scheduled"] = "scheduled",
        ["searchfolders"] = "searchfolders",
        ["sentitems"] = "sentitems",
        ["serverfailures"] = "serverfailures",
        ["syncissues"] = "syncissues"
    };

    public async Task<string> ListAsync(string accessToken, DateTimeOffset? startUtc, DateTimeOffset? endUtc, string readStatus, int top, bool includeBody = false, string folder = null)
    {
        using var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        if (includeBody)
        {
            client.DefaultRequestHeaders.Add("Prefer", "outlook.body-content-type=\"text\"");
        }
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var messages = await ListMessagesAsync(client, startUtc, endUtc, top, readStatus, includeBody: includeBody, folder: folder);

        var output = new JsonObject
        {
            ["count"] = messages.Count,
            ["query"] = new JsonObject
            {
                ["start_utc"] = startUtc?.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture),
                ["end_utc"] = endUtc?.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture),
                ["read_status"] = readStatus,
                ["include_body"] = includeBody,
                ["folder"] = folder
            },
            ["messages"] = new JsonArray(messages.Select(m => m.DeepClone()).ToArray())
        };

        return output.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    public async Task<string> SearchAsync(
        string accessToken,
        DateTimeOffset? startUtc,
        DateTimeOffset? endUtc,
        string readStatus,
        int top,
        string senderQuery,
        string subjectQuery,
        bool includeBody = false)
    {
        using var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        if (includeBody)
        {
            client.DefaultRequestHeaders.Add("Prefer", "outlook.body-content-type=\"text\"");
        }
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var messages = await ListMessagesAsync(client, startUtc, endUtc, top, readStatus, message =>
            MatchesSearch(message, senderQuery, subjectQuery), includeBody: includeBody);

        var output = new JsonObject
        {
            ["count"] = messages.Count,
            ["query"] = new JsonObject
            {
                ["start_utc"] = startUtc?.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture),
                ["end_utc"] = endUtc?.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture),
                ["read_status"] = readStatus,
                ["from_contains"] = senderQuery,
                ["subject_contains"] = subjectQuery,
                ["include_body"] = includeBody
            },
            ["messages"] = new JsonArray(messages.Select(m => m.DeepClone()).ToArray())
        };

        return output.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    public async Task<string> ReadByIdAsync(string accessToken, string messageId)
    {
        using var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        client.DefaultRequestHeaders.Add("Prefer", "outlook.body-content-type=\"text\"");
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var queryParameters = new Dictionary<string, string>
        {
            ["$select"] = "id,subject,isRead,receivedDateTime,from,webLink,body"
        };

        var requestUrl = BuildUrl($"{GraphBaseUrl}/me/messages/{Uri.EscapeDataString(messageId)}", queryParameters);

        using var response = await client.GetAsync(requestUrl);
        var responseText = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {responseText}");
        }

        using var doc = JsonDocument.Parse(responseText);
        if (!TryFormatMessage(doc.RootElement, out var message))
        {
            throw new InvalidOperationException("The requested message could not be parsed.");
        }

        return message.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    private static async Task<List<JsonObject>> ListMessagesAsync(
        HttpClient client,
        DateTimeOffset? startUtc,
        DateTimeOffset? endUtc,
        int top,
        string readStatus,
        Func<JsonObject, bool> predicate = null,
        bool includeBody = false,
        string folder = null)
    {
        var targetCount = top;

        var queryParameters = new Dictionary<string, string>
        {
            ["$select"] = BuildSelectClause(includeBody),
            ["$orderby"] = "receivedDateTime desc",
            ["$top"] = GraphPageSize.ToString(CultureInfo.InvariantCulture)
        };

        var filter = BuildFilter(startUtc, endUtc, readStatus);
        if (!string.IsNullOrWhiteSpace(filter))
        {
            queryParameters["$filter"] = filter;
        }

        var messagesBaseUrl = await ResolveMessagesCollectionUrlAsync(client, folder);
        var nextUrl = BuildUrl(messagesBaseUrl, queryParameters);
        var messages = new List<JsonObject>();

        while (!string.IsNullOrWhiteSpace(nextUrl))
        {
            using var response = await client.GetAsync(nextUrl);
            var responseText = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {responseText}");
            }

            using var doc = JsonDocument.Parse(responseText);
            var root = doc.RootElement;

            if (root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in valueElement.EnumerateArray())
                {
                    if (!TryFormatMessage(item, includeBody, out var formattedMessage))
                    {
                        continue;
                    }

                    if (predicate is not null && !predicate(formattedMessage))
                    {
                        continue;
                    }

                    messages.Add(formattedMessage);

                    if (targetCount > 0 && messages.Count >= targetCount)
                    {
                        return messages;
                    }
                }
            }

            nextUrl = root.TryGetProperty("@odata.nextLink", out var nextLinkElement)
                ? nextLinkElement.GetString()
                : null;
        }

        return messages;
    }

    private static string BuildFilter(DateTimeOffset? startUtc, DateTimeOffset? endUtc, string readStatus)
    {
        var filters = new List<string>();

        if (startUtc.HasValue)
        {
            filters.Add($"receivedDateTime ge {startUtc.Value:yyyy-MM-ddTHH:mm:ssZ}");
        }

        if (endUtc.HasValue)
        {
            filters.Add($"receivedDateTime le {endUtc.Value:yyyy-MM-ddTHH:mm:ssZ}");
        }

        if (string.Equals(readStatus, "read", StringComparison.OrdinalIgnoreCase))
        {
            filters.Add("isRead eq true");
        }
        else if (string.Equals(readStatus, "unread", StringComparison.OrdinalIgnoreCase))
        {
            filters.Add("isRead eq false");
        }

        return filters.Count > 0 ? string.Join(" and ", filters) : string.Empty;
    }

    private static async Task<string> ResolveMessagesCollectionUrlAsync(HttpClient client, string folder)
    {
        if (string.IsNullOrWhiteSpace(folder))
        {
            return $"{GraphBaseUrl}/me/messages";
        }

        var trimmedFolder = folder.Trim();

        if (WellKnownMailFolders.TryGetValue(trimmedFolder, out var wellKnownFolder))
        {
            return $"{GraphBaseUrl}/me/mailFolders/{wellKnownFolder}/messages";
        }

        var folderId = await ResolveFolderIdByDisplayNameAsync(client, trimmedFolder);
        return $"{GraphBaseUrl}/me/mailFolders/{Uri.EscapeDataString(folderId)}/messages";
    }

    private static async Task<string> ResolveFolderIdByDisplayNameAsync(HttpClient client, string folderDisplayName)
    {
        var folders = await GetAllMailFoldersAsync(client);

        var exactMatches = folders
            .Where(folder => string.Equals(folder.DisplayName, folderDisplayName, StringComparison.Ordinal))
            .ToList();

        if (exactMatches.Count == 1)
        {
            return exactMatches[0].Id;
        }

        if (exactMatches.Count > 1)
        {
            throw new InvalidOperationException($"Multiple mail folders were found with the exact name '{folderDisplayName}'. Please use a well-known folder name when possible.");
        }

        var caseInsensitiveMatches = folders
            .Where(folder => string.Equals(folder.DisplayName, folderDisplayName, StringComparison.OrdinalIgnoreCase))
            .Select(folder => folder.DisplayName)
            .Distinct(StringComparer.Ordinal)
            .ToList();

        if (caseInsensitiveMatches.Count == 1)
        {
            throw new InvalidOperationException($"Folder '{folderDisplayName}' was not found with exact casing. Did you mean '{caseInsensitiveMatches[0]}'?");
        }

        if (caseInsensitiveMatches.Count > 1)
        {
            throw new InvalidOperationException($"Folder '{folderDisplayName}' is ambiguous because multiple folders differ only by casing.");
        }

        throw new InvalidOperationException($"Folder '{folderDisplayName}' was not found.");
    }

    private static async Task<List<MailFolderInfo>> GetAllMailFoldersAsync(HttpClient client)
    {
        var discoveredFolders = new List<MailFolderInfo>();
        var pendingCollectionUrls = new Queue<string>();
        pendingCollectionUrls.Enqueue($"{GraphBaseUrl}/me/mailFolders?$select=id,displayName,childFolderCount&$top=100");

        while (pendingCollectionUrls.Count > 0)
        {
            var nextCollectionUrl = pendingCollectionUrls.Dequeue();

            while (!string.IsNullOrWhiteSpace(nextCollectionUrl))
            {
                using var response = await client.GetAsync(nextCollectionUrl);
                var responseText = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {responseText}");
                }

                using var doc = JsonDocument.Parse(responseText);
                var root = doc.RootElement;

                if (root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array)
                {
                    foreach (var folderElement in valueElement.EnumerateArray())
                    {
                        var id = GetStringOrNull(folderElement, "id");
                        if (string.IsNullOrWhiteSpace(id))
                        {
                            continue;
                        }

                        var displayName = GetStringOrNull(folderElement, "displayName");
                        var childFolderCount = GetInt32OrDefault(folderElement, "childFolderCount");

                        discoveredFolders.Add(new MailFolderInfo(id, displayName));

                        if (childFolderCount > 0)
                        {
                            pendingCollectionUrls.Enqueue($"{GraphBaseUrl}/me/mailFolders/{Uri.EscapeDataString(id)}/childFolders?$select=id,displayName,childFolderCount&$top=100");
                        }
                    }
                }

                nextCollectionUrl = root.TryGetProperty("@odata.nextLink", out var nextLinkElement)
                    ? nextLinkElement.GetString()
                    : null;
            }
        }

        return discoveredFolders;
    }

    private static bool MatchesSearch(JsonObject message, string senderQuery, string subjectQuery)
    {
        if (!ContainsText(GetSenderSearchText(message), senderQuery))
        {
            return false;
        }

        return ContainsText(message["subject"]?.GetValue<string>(), subjectQuery);
    }

    private static string GetSenderSearchText(JsonObject message)
    {
        if (message["from"] is not JsonObject fromObject)
        {
            return null;
        }

        var senderName = fromObject["name"]?.GetValue<string>();
        var senderAddress = fromObject["address"]?.GetValue<string>();
        return string.Join(" ", new[] { senderName, senderAddress }.Where(static value => !string.IsNullOrWhiteSpace(value)));
    }

    private static bool ContainsText(string source, string query)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            return true;
        }

        return !string.IsNullOrWhiteSpace(source) &&
               source.Contains(query, StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryFormatMessage(JsonElement message, out JsonObject formattedMessage)
    {
        return TryFormatMessage(message, true, out formattedMessage);
    }

    private static bool TryFormatMessage(JsonElement message, bool includeBody, out JsonObject formattedMessage)
    {
        string senderName = null;
        string senderAddress = null;

        if (message.TryGetProperty("from", out var fromElement) &&
            fromElement.ValueKind == JsonValueKind.Object &&
            fromElement.TryGetProperty("emailAddress", out var emailAddressElement) &&
            emailAddressElement.ValueKind == JsonValueKind.Object)
        {
            if (emailAddressElement.TryGetProperty("name", out var nameElement))
            {
                senderName = nameElement.GetString();
            }

            if (emailAddressElement.TryGetProperty("address", out var addressElement))
            {
                senderAddress = addressElement.GetString();
            }
        }

        formattedMessage = new JsonObject
        {
            ["id"] = GetStringOrNull(message, "id"),
            ["subject"] = GetStringOrNull(message, "subject"),
            ["isRead"] = GetBooleanOrNull(message, "isRead"),
            ["receivedDateTime"] = GetStringOrNull(message, "receivedDateTime"),
            ["webLink"] = GetStringOrNull(message, "webLink"),
            ["from"] = new JsonObject
            {
                ["name"] = senderName,
                ["address"] = senderAddress
            }
        };

        if (includeBody && message.TryGetProperty("body", out var bodyElement))
        {
            formattedMessage["body"] = GetStringOrNull(bodyElement, "content");
        }

        return true;
    }

    private static string BuildSelectClause(bool includeBody)
    {
        return includeBody
            ? "id,subject,isRead,receivedDateTime,from,webLink,body"
            : "id,subject,isRead,receivedDateTime,from,webLink";
    }

    private static string BuildUrl(string baseUrl, Dictionary<string, string> queryParameters)
    {
        var query = string.Join("&", queryParameters.Select(kvp =>
            $"{Uri.EscapeDataString(kvp.Key)}={Uri.EscapeDataString(kvp.Value)}"));

        return $"{baseUrl}?{query}";
    }

    private static string GetStringOrNull(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var value) && value.ValueKind != JsonValueKind.Null
            ? value.GetString()
            : null;
    }

    private static bool? GetBooleanOrNull(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value))
        {
            return null;
        }

        if (value.ValueKind is JsonValueKind.True or JsonValueKind.False)
        {
            return value.GetBoolean();
        }

        return null;
    }

    private static int GetInt32OrDefault(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value) || value.ValueKind != JsonValueKind.Number)
        {
            return 0;
        }

        return value.TryGetInt32(out var number) ? number : 0;
    }

    private sealed record MailFolderInfo(string Id, string DisplayName);
}
