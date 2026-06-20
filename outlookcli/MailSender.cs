using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace outlookcli;

internal sealed class MailSender
{
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    public async Task SendAsync(string accessToken, string to, string subject, string body, string contentType)
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
            ["message"] = new JsonObject
            {
                ["subject"] = subject,
                ["body"] = new JsonObject
                {
                    ["contentType"] = contentType,
                    ["content"] = body
                },
                ["toRecipients"] = recipients
            },
            ["saveToSentItems"] = true
        };

        var json = payload.ToJsonString();
        using var content = new StringContent(json, Encoding.UTF8, "application/json");
        using var response = await client.PostAsync($"{GraphBaseUrl}/me/sendMail", content);

        if (!response.IsSuccessStatusCode)
        {
            var responseText = await response.Content.ReadAsStringAsync();
            throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {responseText}");
        }
    }
}
