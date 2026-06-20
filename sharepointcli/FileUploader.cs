using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace sharepointcli;

internal sealed class FileUploader
{
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    public async Task<int> UploadAsync(string token, string siteUrl, string libraryName, string libraryUrl, string driveId, string filePath, string text, string fileName)
    {
        var content = !string.IsNullOrWhiteSpace(filePath)
            ? await File.ReadAllBytesAsync(Path.GetFullPath(filePath))
            : Encoding.UTF8.GetBytes(text ?? string.Empty);

        var targetFileName = !string.IsNullOrWhiteSpace(fileName)
            ? fileName.Trim()
            : Path.GetFileName(filePath);

        if (string.IsNullOrWhiteSpace(targetFileName))
        {
            throw new ArgumentException("Unable to determine target file name.");
        }

        using var client = CreateClient(token);
        var resolvedDriveId = driveId;

        if (string.IsNullOrWhiteSpace(resolvedDriveId))
        {
            var site = await ResolveSiteAsync(client, siteUrl);
            var drive = await ResolveLibraryAsync(client, GetString(site, "id") ?? throw new InvalidOperationException("Missing site id."), libraryName, libraryUrl);
            resolvedDriveId = GetString(drive, "id") ?? throw new InvalidOperationException("Missing drive id.");
        }

        var uploadUrl = $"{GraphBaseUrl}/drives/{Uri.EscapeDataString(resolvedDriveId)}/root:/{Uri.EscapeDataString(targetFileName)}:/content";
        using var body = new ByteArrayContent(content);
        body.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

        using var response = await client.PutAsync(uploadUrl, body);
        var responseContent = await response.Content.ReadAsStringAsync();
        EnsureSuccess(response, responseContent, "Failed to upload file to SharePoint document library");

        using var doc = JsonDocument.Parse(responseContent);
        var root = doc.RootElement;

        var output = new JsonObject
        {
            ["name"] = GetString(root, "name"),
            ["id"] = GetString(root, "id"),
            ["webUrl"] = GetString(root, "webUrl"),
            ["size"] = GetLong(root, "size"),
            ["createdBy"] = GetCreatedByDisplayName(root),
            ["lastModifiedDateTime"] = GetString(root, "lastModifiedDateTime")
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

    private async Task<JsonElement> ResolveLibraryAsync(HttpClient client, string siteId, string libraryName, string libraryUrl)
    {
        using var response = await client.GetAsync($"{GraphBaseUrl}/sites/{Uri.EscapeDataString(siteId)}/drives?$select=id,name,webUrl");
        var content = await response.Content.ReadAsStringAsync();
        EnsureSuccess(response, content, "Failed to list document libraries for site");

        using var doc = JsonDocument.Parse(content);
        if (!doc.RootElement.TryGetProperty("value", out var drives) || drives.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("Graph did not return any document libraries.");
        }

        foreach (var drive in drives.EnumerateArray())
        {
            if (!string.IsNullOrWhiteSpace(libraryName) && string.Equals(GetString(drive, "name"), libraryName, StringComparison.OrdinalIgnoreCase))
            {
                return drive.Clone();
            }

            if (!string.IsNullOrWhiteSpace(libraryUrl) && string.Equals(GetString(drive, "webUrl"), libraryUrl, StringComparison.OrdinalIgnoreCase))
            {
                return drive.Clone();
            }
        }

        throw new InvalidOperationException(!string.IsNullOrWhiteSpace(libraryName)
            ? $"Document library not found by name: {libraryName}"
            : $"Document library not found by URL: {libraryUrl}");
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

    private static string GetCreatedByDisplayName(JsonElement root)
    {
        if (!root.TryGetProperty("createdBy", out var createdBy) || createdBy.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        if (createdBy.TryGetProperty("user", out var user) && user.ValueKind == JsonValueKind.Object)
        {
            var displayName = GetString(user, "displayName");
            if (!string.IsNullOrWhiteSpace(displayName))
            {
                return displayName;
            }
        }

        if (createdBy.TryGetProperty("application", out var application) && application.ValueKind == JsonValueKind.Object)
        {
            return GetString(application, "displayName");
        }

        return null;
    }

    private static void EnsureSuccess(HttpResponseMessage response, string content, string message)
    {
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"{message}. Graph API error {(int)response.StatusCode}: {content}");
        }
    }
}