using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace teamscli;

internal sealed class TeamsMessageSender
{
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    public async Task SendByUpnAsync(string graphToken, string chatToken, string upn, string message, string contentType)
    {
        var me = await GetMeAsync(graphToken);
        var target = await GetUserByUpnAsync(graphToken, upn);
        var chatId = await CreateOneOnOneChatAsync(chatToken, me.Id, target.Id);
        await SendChatMessageAsync(graphToken, chatId, message, contentType);

        Console.WriteLine($"Teams message sent to {target.UserPrincipalName ?? upn} successfully.");
    }

    public async Task SendByIdAsync(string graphToken, string chatToken, string userId, string message, string contentType)
    {
        var me = await GetMeAsync(graphToken);
        var chatId = await CreateOneOnOneChatAsync(chatToken, me.Id, userId);
        await SendChatMessageAsync(graphToken, chatId, message, contentType);

        Console.WriteLine($"Teams message sent to user id {userId} successfully.");
    }

    public async Task SendToChatAsync(string token, string chatId, string message, string contentType)
    {
        await SendChatMessageAsync(token, chatId, message, contentType);

        Console.WriteLine($"Teams message sent to chat id {chatId} successfully.");
    }

    public async Task SendToChannelAsync(string token, string teamId, string channelId, string message, string contentType)
    {
        using var client = CreateClient(token);
        var payload = new JsonObject
        {
            ["body"] = new JsonObject
            {
                ["contentType"] = contentType,
                ["content"] = message
            }
        };

        var json = payload.ToJsonString();

        using var body = new StringContent(json, Encoding.UTF8, "application/json");
        using var response = await client.PostAsync($"{GraphBaseUrl}/teams/{Uri.EscapeDataString(teamId)}/channels/{Uri.EscapeDataString(channelId)}/messages", body);
        var content = await response.Content.ReadAsStringAsync();

        EnsureSuccess(response, content, "Failed to send channel message");

        Console.WriteLine($"Teams message sent to channel id {channelId} in team id {teamId} successfully.");
    }

    private async Task<(string Id, string UserPrincipalName)> GetMeAsync(string token)
    {
        using var client = CreateClient(token);
        using var response = await client.GetAsync($"{GraphBaseUrl}/me?$select=id,userPrincipalName");
        var content = await response.Content.ReadAsStringAsync();

        EnsureSuccess(response, content, "Failed to get /me");

        using var doc = JsonDocument.Parse(content);
        var root = doc.RootElement;
        return (
            Id: root.GetProperty("id").GetString() ?? throw new InvalidOperationException("Missing me.id"),
            UserPrincipalName: root.TryGetProperty("userPrincipalName", out var upnElement) ? upnElement.GetString() : null
        );
    }

    private async Task<(string Id, string UserPrincipalName)> GetUserByUpnAsync(string token, string upn)
    {
        using var client = CreateClient(token);
        using var response = await client.GetAsync($"{GraphBaseUrl}/users/{Uri.EscapeDataString(upn)}?$select=id,userPrincipalName,displayName");
        var content = await response.Content.ReadAsStringAsync();

        EnsureSuccess(response, content, $"Failed to get user by UPN '{upn}'");

        using var doc = JsonDocument.Parse(content);
        var root = doc.RootElement;
        return (
            Id: root.GetProperty("id").GetString() ?? throw new InvalidOperationException("Missing target user id"),
            UserPrincipalName: root.TryGetProperty("userPrincipalName", out var upnElement) ? upnElement.GetString() : null
        );
    }

    private async Task<string> CreateOneOnOneChatAsync(string token, string meId, string targetUserId)
    {
        using var client = CreateClient(token);
        var members = new JsonArray();
        members.Add((JsonNode)new JsonObject
        {
            ["@odata.type"] = "#microsoft.graph.aadUserConversationMember",
            ["roles"] = new JsonArray("owner"),
            ["user@odata.bind"] = $"https://graph.microsoft.com/v1.0/users('{meId}')"
        });
        members.Add((JsonNode)new JsonObject
        {
            ["@odata.type"] = "#microsoft.graph.aadUserConversationMember",
            ["roles"] = new JsonArray("owner"),
            ["user@odata.bind"] = $"https://graph.microsoft.com/v1.0/users('{targetUserId}')"
        });

        var payload = new JsonObject
        {
            ["chatType"] = "oneOnOne",
            ["members"] = members
        };

        var json = payload.ToJsonString();

        using var body = new StringContent(json, Encoding.UTF8, "application/json");
        using var response = await client.PostAsync($"{GraphBaseUrl}/chats", body);
        var content = await response.Content.ReadAsStringAsync();

        EnsureSuccess(response, content, "Failed to create one-on-one chat");

        using var doc = JsonDocument.Parse(content);
        var id = doc.RootElement.GetProperty("id").GetString();
        return id ?? throw new InvalidOperationException("Graph did not return chat id.");
    }

    private async Task SendChatMessageAsync(string token, string chatId, string message, string contentType)
    {
        using var client = CreateClient(token);
        var payload = new JsonObject
        {
            ["body"] = new JsonObject
            {
                ["contentType"] = contentType,
                ["content"] = message
            }
        };

        var json = payload.ToJsonString();

        using var body = new StringContent(json, Encoding.UTF8, "application/json");
        using var response = await client.PostAsync($"{GraphBaseUrl}/chats/{chatId}/messages", body);
        var content = await response.Content.ReadAsStringAsync();

        EnsureSuccess(response, content, "Failed to send chat message");
    }

    private static HttpClient CreateClient(string token)
    {
        var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        return client;
    }

    private static void EnsureSuccess(HttpResponseMessage response, string content, string message)
    {
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"{message}: {(int)response.StatusCode} {content}");
        }
    }
}
