using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace teamscli;

internal sealed class MeetingTranscriptReader
{
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    public async Task<string> ReadTranscriptAsync(string outlooktoken,string salesInsightsWebAppToken, string joinWebUrl)
    {
        using var client = CreateClient(outlooktoken);

        // Step 1: Resolve the online meeting by joinWebUrl
        var meeting = await GetOnlineMeetingByJoinWebUrlAsync(client, joinWebUrl);
        var meetingId = GetString(meeting, "id")
            ?? throw new InvalidOperationException("Unable to resolve meeting id from joinWebUrl.");

        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", salesInsightsWebAppToken);
        // Step 2: List transcripts for the meeting
        var transcripts = await ListTranscriptsAsync(client, meetingId);

        // Step 3: Fetch transcript content for each transcript
        var transcriptItems = new JsonArray();
        foreach (var transcript in transcripts)
        {
            var transcriptId = GetString(transcript, "id");
            var createdDateTime = GetString(transcript, "createdDateTime");
            var transcriptContentUrl = GetString(transcript, "transcriptContentUrl");

            string content = null;
            if (!string.IsNullOrWhiteSpace(transcriptContentUrl))
            {
                content = await FetchTranscriptContentAsync(client, transcriptContentUrl);
            }

            transcriptItems.Add(new JsonObject
            {
                ["id"] = transcriptId,
                ["createdDateTime"] = createdDateTime,
                ["content"] = content
            });
        }

        var output = new JsonObject
        {
            ["meetingId"] = meetingId,
            ["joinWebUrl"] = GetString(meeting, "joinWebUrl"),
            ["subject"] = GetString(meeting, "subject"),
            ["createdDateTime"] = GetString(meeting, "createdDateTime"),
            ["transcriptCount"] = transcriptItems.Count,
            ["transcripts"] = transcriptItems
        };

        return output.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    private static async Task<JsonObject> GetOnlineMeetingByJoinWebUrlAsync(HttpClient client, string joinWebUrl)
    {
        var encodedUrl = Uri.EscapeDataString(joinWebUrl);
        var url = $"{GraphBaseUrl}/me/onlineMeetings?$filter=JoinWebUrl eq '{encodedUrl}'";

        using var response = await client.GetAsync(url);
        var content = await response.Content.ReadAsStringAsync();
        EnsureSuccess(response, content, "Failed to get online meeting by joinWebUrl");

        using var doc = JsonDocument.Parse(content);
        var root = doc.RootElement;

        if (root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in valueElement.EnumerateArray())
            {
                var meeting = JsonObject.Create(item.Clone(), new JsonNodeOptions());
                if (meeting != null)
                {
                    return meeting;
                }
            }
        }

        throw new InvalidOperationException($"No online meeting found for joinWebUrl: {joinWebUrl}");
    }

    private static async Task<List<JsonObject>> ListTranscriptsAsync(HttpClient client, string meetingId)
    {
        var transcripts = new List<JsonObject>();
        var url = $"{GraphBaseUrl}/me/onlineMeetings/{Uri.EscapeDataString(meetingId)}/transcripts";

        while (!string.IsNullOrWhiteSpace(url))
        {
            using var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();
            EnsureSuccess(response, content, "Failed to list meeting transcripts");

            using var doc = JsonDocument.Parse(content);
            var root = doc.RootElement;

            if (root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in valueElement.EnumerateArray())
                {
                    var transcript = JsonObject.Create(item.Clone(), new JsonNodeOptions());
                    if (transcript != null)
                    {
                        transcripts.Add(transcript);
                    }
                }
            }

            url = root.TryGetProperty("@odata.nextLink", out var nextElement)
                ? nextElement.GetString()
                : null;
        }

        return transcripts;
    }

    private static async Task<string> FetchTranscriptContentAsync(HttpClient client, string transcriptContentUrl)
    {
        try
        {
            using var response = await client.GetAsync(transcriptContentUrl);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }
        catch (Exception ex)
        {
            SystemEventLogger.Error($"Failed to fetch transcript content from {transcriptContentUrl}: {ex.Message}");
            return null;
        }
    }

    private static HttpClient CreateClient(string token)
    {
        var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        return client;
    }

    private static string GetString(JsonObject obj, string propertyName)
    {
        if (obj.TryGetPropertyValue(propertyName, out var node) && node != null)
        {
            return node.GetValue<string>();
        }

        return null;
    }

    private static void EnsureSuccess(HttpResponseMessage response, string content, string message)
    {
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"{message}: {(int)response.StatusCode} {content}");
        }
    }
}
