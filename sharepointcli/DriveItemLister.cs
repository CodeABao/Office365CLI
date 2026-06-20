using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace sharepointcli;

internal sealed class DriveItemLister
{
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";
    private const int GraphPageSize = 50;

    public async Task<int> ListAsync(string token, string driveId, string itemType, int top, string orderBy, bool desc)
    {
        using var client = CreateClient(token);

        var allItems = await FetchRootChildrenUpToTopAsync(client, driveId, itemType, top);
        var sorted = SortItems(allItems, orderBy, desc);
        var limited = sorted.ToArray();

        var results = new JsonArray();
        foreach (var item in limited)
        {
            var isFolder = item.TryGetProperty("folder", out var folderElement) && folderElement.ValueKind == JsonValueKind.Object;
            var isFile = item.TryGetProperty("file", out var fileElement) && fileElement.ValueKind == JsonValueKind.Object;

            results.Add(new JsonObject
            {
                ["id"] = GetString(item, "id"),
                ["name"] = GetString(item, "name"),
                ["itemType"] = isFolder ? "folder" : (isFile ? "file" : "unknown"),
                ["size"] = GetLong(item, "size"),
                ["webUrl"] = GetString(item, "webUrl"),
                ["createdDateTime"] = GetString(item, "createdDateTime"),
                ["lastModifiedDateTime"] = GetString(item, "lastModifiedDateTime")
            });
        }

        var output = new JsonObject
        {
            ["query"] = new JsonObject
            {
                ["driveId"] = driveId,
                ["itemType"] = itemType,
                ["top"] = top,
                ["orderBy"] = orderBy,
                ["desc"] = desc
            },
            ["count"] = results.Count,
            ["results"] = results
        };

        Console.WriteLine(output.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        return 0;
    }

    private async Task<List<JsonElement>> FetchRootChildrenUpToTopAsync(HttpClient client, string driveId, string itemType, int top)
    {
        var items = new List<JsonElement>();
        var nextUrl = $"{GraphBaseUrl}/drives/{Uri.EscapeDataString(driveId)}/root/children?$select=id,name,size,webUrl,createdDateTime,lastModifiedDateTime,file,folder&$top={GraphPageSize}";

        while (!string.IsNullOrWhiteSpace(nextUrl))
        {
            using var response = await client.GetAsync(nextUrl);
            var content = await response.Content.ReadAsStringAsync();
            EnsureSuccess(response, content, "Failed to list drive items");

            using var doc = JsonDocument.Parse(content);
            var root = doc.RootElement;

            if (root.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in value.EnumerateArray())
                {
                    if (!IsTypeMatch(item, itemType))
                    {
                        continue;
                    }

                    items.Add(item.Clone());

                    if (top > 0 && items.Count >= top)
                    {
                        return items;
                    }
                }
            }

            nextUrl = root.TryGetProperty("@odata.nextLink", out var nextLink) && nextLink.ValueKind == JsonValueKind.String
                ? nextLink.GetString()
                : null;
        }

        return items;
    }

    private static bool IsTypeMatch(JsonElement item, string itemType)
    {
        return itemType.ToLowerInvariant() switch
        {
            "folder" => item.TryGetProperty("folder", out var folder) && folder.ValueKind == JsonValueKind.Object,
            "file" => item.TryGetProperty("file", out var file) && file.ValueKind == JsonValueKind.Object,
            _ => true
        };
    }

    private static IOrderedEnumerable<JsonElement> SortItems(IEnumerable<JsonElement> items, string orderBy, bool desc)
    {
        Func<JsonElement, object> selector = orderBy.ToLowerInvariant() switch
        {
            "size" => i => GetLong(i, "size"),
            "createddatetime" => i => GetString(i, "createdDateTime") ?? string.Empty,
            "lastmodifieddatetime" => i => GetString(i, "lastModifiedDateTime") ?? string.Empty,
            _ => i => GetString(i, "name") ?? string.Empty
        };

        return desc
            ? items.OrderByDescending(selector).ThenBy(i => GetString(i, "name") ?? string.Empty)
            : items.OrderBy(selector).ThenBy(i => GetString(i, "name") ?? string.Empty);
    }

    private static HttpClient CreateClient(string token)
    {
        var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        return client;
    }

    private static string GetString(JsonElement root, string propertyName)
    {
        return root.TryGetProperty(propertyName, out var value) && value.ValueKind != JsonValueKind.Null
            ? value.GetString()
            : null;
    }

    private static long GetLong(JsonElement root, string propertyName)
    {
        return root.TryGetProperty(propertyName, out var value) && value.TryGetInt64(out var parsed)
            ? parsed
            : 0;
    }

    private static void EnsureSuccess(HttpResponseMessage response, string content, string message)
    {
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"{message}. Graph API error {(int)response.StatusCode}: {content}");
        }
    }
}
