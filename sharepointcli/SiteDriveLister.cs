using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace sharepointcli;

internal sealed class SiteDriveLister
{
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    public async Task<int> ListAsync(string token, string siteUrl)
    {
        using var client = CreateClient(token);
        var site = await ResolveSiteAsync(client, siteUrl);
        var siteId = GetString(site, "id") ?? throw new InvalidOperationException("Missing site id.");

        using var response = await client.GetAsync($"{GraphBaseUrl}/sites/{Uri.EscapeDataString(siteId)}/drives?$select=id,name,webUrl,driveType,createdDateTime,lastModifiedDateTime");
        var content = await response.Content.ReadAsStringAsync();
        EnsureSuccess(response, content, "Failed to list drives for site");

        using var doc = JsonDocument.Parse(content);
        var drives = new JsonArray();
        if (doc.RootElement.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array)
        {
            foreach (var drive in value.EnumerateArray())
            {
                drives.Add(new JsonObject
                {
                    ["id"] = GetString(drive, "id"),
                    ["name"] = GetString(drive, "name"),
                    ["driveType"] = GetString(drive, "driveType"),
                    ["webUrl"] = GetString(drive, "webUrl"),
                    ["createdDateTime"] = GetString(drive, "createdDateTime"),
                    ["lastModifiedDateTime"] = GetString(drive, "lastModifiedDateTime")
                });
            }
        }

        var output = new JsonObject
        {
            ["site"] = new JsonObject
            {
                ["id"] = GetString(site, "id"),
                ["displayName"] = GetString(site, "displayName"),
                ["webUrl"] = GetString(site, "webUrl")
            },
            ["count"] = drives.Count,
            ["drives"] = drives
        };

        Console.WriteLine(output.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        return 0;
    }

    private async Task<JsonElement> ResolveSiteAsync(HttpClient client, string siteUrl)
    {
        var uri = new Uri(siteUrl);
        var sitePath = uri.AbsolutePath.TrimEnd('/');
        if (string.IsNullOrWhiteSpace(sitePath))
        {
            throw new ArgumentException($"Invalid site URL: {siteUrl}");
        }

        var requestUrl = $"{GraphBaseUrl}/sites/{Uri.EscapeDataString(uri.Host)}:{sitePath}?$select=id,displayName,webUrl";
        using var response = await client.GetAsync(requestUrl);
        var content = await response.Content.ReadAsStringAsync();
        EnsureSuccess(response, content, "Failed to resolve SharePoint site");

        using var doc = JsonDocument.Parse(content);
        return doc.RootElement.Clone();
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

    private static void EnsureSuccess(HttpResponseMessage response, string content, string message)
    {
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"{message}. Graph API error {(int)response.StatusCode}: {content}");
        }
    }
}
