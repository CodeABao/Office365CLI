using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace outlookcli;

internal sealed class MailReplier
{
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    public async Task ReplyAsync(string accessToken, string messageId, string body, string contentType)
    {
        using var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

        using var createResponse = await client.PostAsync($"{GraphBaseUrl}/me/messages/{Uri.EscapeDataString(messageId)}/createReply", content: null);
        var createResponseText = await createResponse.Content.ReadAsStringAsync();

        if (!createResponse.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Graph API error {(int)createResponse.StatusCode}: {createResponseText}");
        }

        var node = JsonNode.Parse(createResponseText) as JsonObject;
        var draftId = node?["id"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(draftId))
        {
            throw new InvalidOperationException("Graph API did not return a draft id for the reply message.");
        }

        var existingBodyNode = node?["body"] as JsonObject;
        var existingBody = existingBodyNode?["content"]?.GetValue<string>() ?? string.Empty;
        var existingContentType = existingBodyNode?["contentType"]?.GetValue<string>() ?? contentType;
        var mergedContentType = NormalizeBodyContentType(existingContentType, contentType);
        var mergedBody = BuildReplyBody(body, existingBody, mergedContentType);

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

        using var response = await client.PostAsync($"{GraphBaseUrl}/me/messages/{Uri.EscapeDataString(draftId)}/send", content: null);

        if (!response.IsSuccessStatusCode)
        {
            var responseText = await response.Content.ReadAsStringAsync();
            throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {responseText}");
        }
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

    private static string BuildReplyBody(string userBody, string originalBody, string contentType)
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