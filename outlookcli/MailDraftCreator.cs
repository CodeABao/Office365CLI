using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace outlookcli;

internal sealed class MailDraftCreator
{
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    public async Task<string> CreateDraftAsync(string accessToken, string to, string subject, string body, string contentType)
    {
        using var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

        var recipients = new JsonArray();
        recipients.Add((JsonNode)new JsonObject
        {
            ["emailAddress"] = new JsonObject
            {
                ["address"] = to
            }
        });

        var payload = new JsonObject
        {
            ["subject"] = subject,
            ["body"] = new JsonObject
            {
                ["contentType"] = contentType,
                ["content"] = body
            },
            ["toRecipients"] = recipients
        };

        var json = payload.ToJsonString();
        using var content = new StringContent(json, Encoding.UTF8, "application/json");
        using var response = await client.PostAsync($"{GraphBaseUrl}/me/messages", content);
        var responseText = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {responseText}");
        }

        var node = JsonNode.Parse(responseText);
        return node?.ToJsonString(new JsonSerializerOptions { WriteIndented = true }) ?? responseText;
    }

    public async Task<string> CreateReplyDraftAsync(string accessToken, string messageId, string body, string contentType)
    {
        using var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

        using var response = await client.PostAsync($"{GraphBaseUrl}/me/messages/{Uri.EscapeDataString(messageId)}/createReply", content: null);
        var responseText = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {responseText}");
        }

        var node = JsonNode.Parse(responseText) as JsonObject;
        var draftId = node?["id"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(draftId))
        {
            return node?.ToJsonString(new JsonSerializerOptions { WriteIndented = true }) ?? responseText;
        }

        var existingBodyNode = node?["body"] as JsonObject;
        var existingBody = existingBodyNode?["content"]?.GetValue<string>() ?? string.Empty;
        var existingContentType = existingBodyNode?["contentType"]?.GetValue<string>() ?? contentType;
        var mergedContentType = NormalizeBodyContentType(existingContentType, contentType);
        var mergedBody = BuildReplyDraftBody(body, existingBody, mergedContentType);

        var updatePayload = new JsonObject
        {
            ["body"] = new JsonObject
            {
                ["contentType"] = mergedContentType,
                ["content"] = mergedBody
            }
        };

        using var patchRequest = new HttpRequestMessage(new HttpMethod("PATCH"), $"{GraphBaseUrl}/me/messages/{Uri.EscapeDataString(draftId)}")
        {
            Content = new StringContent(updatePayload.ToJsonString(), Encoding.UTF8, "application/json")
        };
        using var patchResponse = await client.SendAsync(patchRequest);
        var patchResponseText = await patchResponse.Content.ReadAsStringAsync();

        if (!patchResponse.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Graph API error {(int)patchResponse.StatusCode}: {patchResponseText}");
        }

        node!["body"] = new JsonObject
        {
            ["contentType"] = mergedContentType,
            ["content"] = mergedBody
        };

        return node.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    private static string NormalizeBodyContentType(string existingContentType, string requestedContentType)
    {
        if (!string.IsNullOrWhiteSpace(existingContentType))
        {
            return string.Equals(existingContentType, "html", StringComparison.OrdinalIgnoreCase)
                ? "HTML"
                : string.Equals(existingContentType, "text", StringComparison.OrdinalIgnoreCase)
                    ? "Text"
                    : existingContentType;
        }

        return requestedContentType;
    }

    private static string BuildReplyDraftBody(string userBody, string originalBody, string contentType)
    {
        if (string.IsNullOrEmpty(originalBody))
        {
            return userBody;
        }

        return string.Equals(contentType, "HTML", StringComparison.OrdinalIgnoreCase)
            ? $"{userBody}<br><br>{originalBody}"
            : $"{userBody}{Environment.NewLine}{Environment.NewLine}{originalBody}";
    }
}
