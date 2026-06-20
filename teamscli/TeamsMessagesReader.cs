using System.Text.Json;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;

namespace teamscli;

internal sealed class TeamsMessagesReader
{
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";
    private const int MaxReadMessageTop = 500;
    private const int MaxChatPagesWhenReadFilterProvided = 4;

    public async Task<string> ListMessagesAsync(string token, string upn, string chatId, bool? isRead, int top, DateTimeOffset? startUtc, DateTimeOffset? endUtc)
    {
        using var client = CreateClient(token, preferTextBody: true);

        var me = await GetMeAsync(client);
        JsonObject targetUser = null;
        var targetUserId = string.Empty;
        var resolvedChatId = chatId;

        if (string.IsNullOrWhiteSpace(resolvedChatId) && !string.IsNullOrWhiteSpace(upn))
        {
            targetUser = await GetUserByUpnAsync(client, upn);
            targetUserId = GetString(targetUser, "id") ?? string.Empty;
            var meId = GetString(me, "id") ?? string.Empty;
            resolvedChatId = await CreateOneOnOneChatAsync(client, meId, targetUserId);
        }

        if (string.IsNullOrWhiteSpace(resolvedChatId))
        {
            throw new InvalidOperationException("Unable to resolve chat id for list-messages.");
        }

        var requestedTop = Math.Min(top, MaxReadMessageTop);
        var chat = await GetChatByIdAsync(client, resolvedChatId);
        var matched = new List<JsonObject>();
        var skippedUnknownFutureValue = 0;
        var scannedMessages = 0;

        var chatMessages = await ListRecentChatMessagesAsync(client, resolvedChatId, requestedTop, startUtc, endUtc);
        foreach (var msg in chatMessages)
        {
            scannedMessages++;

            var messageType = GetString(msg, "messageType");
            if (string.Equals(messageType, "unknownFutureValue", StringComparison.OrdinalIgnoreCase))
            {
                skippedUnknownFutureValue++;
                continue;
            }

            if (!MatchesReadMessageFilter(msg, targetUserId))
            {
                continue;
            }

            var formatted = FormatMessage(chat, msg, me);
            if (isRead.HasValue)
            {
                var readState = GetString(formatted, "readState") ?? string.Empty;
                var currentIsRead = string.Equals(readState, "read", StringComparison.OrdinalIgnoreCase);
                if (currentIsRead != isRead.Value)
                {
                    continue;
                }
            }

            matched.Add(formatted);

            if (matched.Count >= requestedTop)
            {
                break;
            }
        }

        if (skippedUnknownFutureValue > 0)
        {
            SystemEventLogger.Info($"Skipped {skippedUnknownFutureValue} Teams messages with messageType=unknownFutureValue.");
        }

        var resultItems = new JsonArray();
        foreach (var item in matched
            .OrderByDescending(x => GetString(x, "createdDateTime") ?? string.Empty)
            .Take(requestedTop))
        {
            var messageNode = new JsonObject();
            foreach (var kvp in item)
            {
                if (kvp.Key == "chat")
                {
                    continue;
                }

                messageNode[kvp.Key] = kvp.Value?.DeepClone();
            }

            resultItems.Add(messageNode);
        }

        var query = new JsonObject
        {
            ["top"] = requestedTop,
            ["chat_count_scanned"] = 1,
            ["message_count_scanned"] = scannedMessages
        };

        query["is_read"] = isRead.HasValue ? isRead.Value : "all";

        if (!string.IsNullOrWhiteSpace(upn))
        {
            query["upn"] = upn;
        }

        if (!string.IsNullOrWhiteSpace(resolvedChatId))
        {
            query["chat_id"] = resolvedChatId;
        }

        if (startUtc.HasValue)
        {
            query["start_utc"] = startUtc.Value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
        }

        if (endUtc.HasValue)
        {
            query["end_utc"] = endUtc.Value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
        }

        var output = new JsonObject
        {
            ["count"] = resultItems.Count,
            ["query"] = query,
            ["me"] = me,
            ["targetUser"] = targetUser?.DeepClone(),
            ["messages"] = resultItems
        };

        return output.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    private static async Task<JsonElement> GetChatByIdAsync(HttpClient client, string chatId)
    {
        var query = new Dictionary<string, string>
        {
            ["$select"] = "id,chatType,topic,lastUpdatedDateTime,lastMessagePreview,webUrl,viewpoint",
            ["$expand"] = "lastMessagePreview"
        };

        using var response = await client.GetAsync(BuildUrl($"{GraphBaseUrl}/me/chats/{Uri.EscapeDataString(chatId)}", query));
        var content = await response.Content.ReadAsStringAsync();
        EnsureSuccess(response, content);

        using var doc = JsonDocument.Parse(content);
        return doc.RootElement.Clone();
    }

    public async Task<string> ListChatsAsync(string token, string name, string chatType, bool? isRead, int top)
    {
        using var client = CreateClient(token);

        var requestedTop = Math.Min(top, MaxReadMessageTop);
        var normalizedName = name?.Trim() ?? string.Empty;
        var maxPageCount = isRead == false ? MaxChatPagesWhenReadFilterProvided : (int?)null;
        var (filteredChats, scannedChatCount, scannedPageCount) = await ListChatsAsync(
            client,
            requestedTop,
            chatType,
            null,
            null,
            normalizedName,
            isRead,
            maxPageCount);

        var resultItems = new JsonArray();
        foreach (var chat in filteredChats)
        {
            var chatReadState = IsChatRead(chat);
            resultItems.Add((JsonNode)new JsonObject
            {
                ["id"] = GetString(chat, "id"),
                ["topic"] = GetString(chat, "topic"),
                ["type"] = GetString(chat, "chatType"),
                ["webUrl"] = GetString(chat, "webUrl"),
                ["lastUpdatedDateTime"] = GetString(chat, "lastUpdatedDateTime"),
                ["lastMessageReadDateTime"] = GetChatLastReadDateTime(chat),
                ["isRead"] = chatReadState.HasValue ? chatReadState.Value : null
            });
        }

        var output = new JsonObject
        {
            ["count"] = resultItems.Count,
            ["query"] = new JsonObject
            {
                ["name"] = string.IsNullOrWhiteSpace(normalizedName) ? null : normalizedName,
                ["top"] = requestedTop,
                ["chat_type"] = chatType,
                ["is_read"] = isRead.HasValue ? isRead.Value : "all",
                ["chat_count_scanned"] = scannedChatCount,
                ["page_count_scanned"] = scannedPageCount,
                ["page_limit"] = maxPageCount
            },
            ["chats"] = resultItems
        };

        return output.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    public async Task<string> ListJoinedTeamsAsync(string token, string name, int top)
    {
        using var client = CreateClient(token);

        var requestedTop = Math.Min(top, MaxReadMessageTop);
        var normalizedName = name?.Trim() ?? string.Empty;
        var (teams, scannedTeamCount) = await ListJoinedTeamsInternalAsync(client, requestedTop, normalizedName);

        var resultItems = new JsonArray();
        foreach (var team in teams.OrderBy(t => GetString(t, "displayName") ?? string.Empty))
        {
            resultItems.Add((JsonNode)new JsonObject
            {
                ["id"] = GetString(team, "id"),
                ["displayName"] = GetString(team, "displayName"),
                ["description"] = GetString(team, "description"),
                ["visibility"] = GetString(team, "visibility"),
                ["isArchived"] = team.TryGetProperty("isArchived", out var isArchived) && isArchived.ValueKind is JsonValueKind.True or JsonValueKind.False
                    ? isArchived.GetBoolean()
                    : null,
                ["webUrl"] = GetString(team, "webUrl")
            });
        }

        var output = new JsonObject
        {
            ["count"] = resultItems.Count,
            ["query"] = new JsonObject
            {
                ["name"] = string.IsNullOrWhiteSpace(normalizedName) ? null : normalizedName,
                ["top"] = requestedTop,
                ["team_count_scanned"] = scannedTeamCount
            },
            ["teams"] = resultItems
        };

        return output.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    public async Task<string> ListTeamChannelsAsync(string token, string teamId, string name)
    {
        using var client = CreateClient(token);

        var normalizedName = name?.Trim() ?? string.Empty;
        var channels = await ListTeamChannelsInternalAsync(client, teamId);

        var filteredChannels = string.IsNullOrWhiteSpace(normalizedName)
            ? channels
            : channels.Where(c =>
                (GetString(c, "displayName") ?? string.Empty).Contains(normalizedName, StringComparison.OrdinalIgnoreCase))
                .ToList();

        var resultItems = new JsonArray();
        foreach (var channel in filteredChannels
            .OrderBy(c => GetString(c, "displayName") ?? string.Empty))
        {
            resultItems.Add((JsonNode)new JsonObject
            {
                ["id"] = GetString(channel, "id"),
                ["displayName"] = GetString(channel, "displayName"),
                ["description"] = GetString(channel, "description"),
                ["membershipType"] = GetString(channel, "membershipType"),
                ["webUrl"] = GetString(channel, "webUrl")
            });
        }

        var output = new JsonObject
        {
            ["count"] = resultItems.Count,
            ["query"] = new JsonObject
            {
                ["team_id"] = teamId,
                ["name"] = string.IsNullOrWhiteSpace(normalizedName) ? null : normalizedName,
                ["channel_count_scanned"] = channels.Count
            },
            ["channels"] = resultItems
        };

        return output.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    public async Task<string> ListChannelMessagesAsync(string token, string teamId, string channelId, int top)
    {
        using var client = CreateClient(token, preferTextBody: true);

        var requestedTop = Math.Min(top, MaxReadMessageTop);
        var messages = await ListChannelMessagesInternalAsync(client, teamId, channelId, requestedTop);
        var skippedUnknownFutureValue = 0;

        var resultItems = new JsonArray();
        foreach (var message in messages
            .OrderByDescending(m => GetString(m, "createdDateTime") ?? string.Empty)
            .Take(requestedTop))
        {
            var messageType = GetString(message, "messageType");
            if (string.Equals(messageType, "unknownFutureValue", StringComparison.OrdinalIgnoreCase))
            {
                skippedUnknownFutureValue++;
                continue;
            }

            var senderUser = message.TryGetProperty("from", out var fromElement) &&
                             fromElement.ValueKind == JsonValueKind.Object &&
                             fromElement.TryGetProperty("user", out var userElement) &&
                             userElement.ValueKind == JsonValueKind.Object
                ? userElement
                : default;

            resultItems.Add((JsonNode)new JsonObject
            {
                ["id"] = GetString(message, "id"),
                ["replyToId"] = GetString(message, "replyToId"),
                ["createdDateTime"] = GetString(message, "createdDateTime"),
                ["lastModifiedDateTime"] = GetString(message, "lastModifiedDateTime"),
                ["webUrl"] = GetString(message, "webUrl"),
                ["from"] = new JsonObject
                {
                    ["id"] = GetString(senderUser, "id"),
                    ["displayName"] = GetString(senderUser, "displayName"),
                    ["userIdentityType"] = GetString(senderUser, "userIdentityType")
                },
                ["body"] = message.TryGetProperty("body", out var bodyElement)
                    ? new JsonObject
                    {
                        ["contentType"] = GetString(bodyElement, "contentType"),
                        ["content"] = GetString(bodyElement, "content")
                    }
                    : null
            });
        }

        if (skippedUnknownFutureValue > 0)
        {
            SystemEventLogger.Info($"Skipped {skippedUnknownFutureValue} channel messages with messageType=unknownFutureValue.");
        }

        var output = new JsonObject
        {
            ["count"] = resultItems.Count,
            ["query"] = new JsonObject
            {
                ["team_id"] = teamId,
                ["channel_id"] = channelId,
                ["top"] = requestedTop,
                ["message_count_scanned"] = messages.Count
            },
            ["messages"] = resultItems
        };

        return output.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    private static async Task<(List<JsonElement> Teams, int ScannedCount)> ListJoinedTeamsInternalAsync(HttpClient client, int maxTeams, string name)
    {
        var url = $"{GraphBaseUrl}/me/joinedTeams";
        var teams = new List<JsonElement>();
        var scannedCount = 0;

        while (!string.IsNullOrWhiteSpace(url) && teams.Count < maxTeams)
        {
            using var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();
            EnsureSuccess(response, content);

            using var doc = JsonDocument.Parse(content);
            var root = doc.RootElement;

            if (root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var team in valueElement.EnumerateArray())
                {
                    scannedCount++;

                    if (!MatchesTeamName(team, name))
                    {
                        continue;
                    }

                    teams.Add(team.Clone());
                    if (teams.Count >= maxTeams)
                    {
                        break;
                    }
                }
            }

            url = root.TryGetProperty("@odata.nextLink", out var nextLink) ? nextLink.GetString() : null;
        }

        return (teams, scannedCount);
    }

    private static bool MatchesTeamName(JsonElement team, string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return true;
        }

        return (GetString(team, "displayName") ?? string.Empty).Contains(name, StringComparison.OrdinalIgnoreCase);
    }

    private static async Task<List<JsonElement>> ListTeamChannelsInternalAsync(HttpClient client, string teamId)
    {
        var query = new Dictionary<string, string>
        {
            ["$select"] = "id,displayName,description,membershipType,webUrl"
        };

        var url = BuildUrl($"{GraphBaseUrl}/teams/{Uri.EscapeDataString(teamId)}/channels", query);
        var channels = new List<JsonElement>();

        while (!string.IsNullOrWhiteSpace(url))
        {
            using var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();
            EnsureSuccess(response, content);

            using var doc = JsonDocument.Parse(content);
            var root = doc.RootElement;
            var pageItems = root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array
                ? valueElement.EnumerateArray().Select(x => x.Clone()).ToList()
                : new List<JsonElement>();

            channels.AddRange(pageItems);
            url = root.TryGetProperty("@odata.nextLink", out var nextLink) ? nextLink.GetString() : null;
        }

        return channels;
    }

    private static async Task<List<JsonElement>> ListChannelMessagesInternalAsync(HttpClient client, string teamId, string channelId, int maxMessages)
    {
        var query = new Dictionary<string, string>
        {
            ["$top"] = "50"
        };

        var url = BuildUrl($"{GraphBaseUrl}/teams/{Uri.EscapeDataString(teamId)}/channels/{Uri.EscapeDataString(channelId)}/messages", query);
        var messages = new List<JsonElement>();

        while (!string.IsNullOrWhiteSpace(url) && messages.Count < maxMessages)
        {
            using var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();
            EnsureSuccess(response, content);

            using var doc = JsonDocument.Parse(content);
            var root = doc.RootElement;
            var pageItems = root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array
                ? valueElement.EnumerateArray().Select(x => x.Clone()).ToList()
                : new List<JsonElement>();

            foreach (var item in pageItems)
            {
                messages.Add(item);
                if (messages.Count >= maxMessages)
                {
                    break;
                }
            }

            url = root.TryGetProperty("@odata.nextLink", out var nextLink) ? nextLink.GetString() : null;
        }

        return messages;
    }

    private static async Task<JsonObject> GetMeAsync(HttpClient client)
    {
        using var response = await client.GetAsync($"{GraphBaseUrl}/me?$select=id,displayName,mail,userPrincipalName");
        var content = await response.Content.ReadAsStringAsync();
        EnsureSuccess(response, content);

        using var doc = JsonDocument.Parse(content);
        var root = doc.RootElement;
        return new JsonObject
        {
            ["id"] = GetString(root, "id"),
            ["displayName"] = GetString(root, "displayName"),
            ["mail"] = GetString(root, "mail"),
            ["userPrincipalName"] = GetString(root, "userPrincipalName")
        };
    }

    private static async Task<JsonObject> GetUserByUpnAsync(HttpClient client, string upn)
    {
        using var response = await client.GetAsync($"{GraphBaseUrl}/users/{Uri.EscapeDataString(upn)}?$select=id,displayName,mail,userPrincipalName");
        var content = await response.Content.ReadAsStringAsync();
        EnsureSuccess(response, content);

        using var doc = JsonDocument.Parse(content);
        var root = doc.RootElement;
        return new JsonObject
        {
            ["id"] = GetString(root, "id"),
            ["displayName"] = GetString(root, "displayName"),
            ["mail"] = GetString(root, "mail"),
            ["userPrincipalName"] = GetString(root, "userPrincipalName")
        };
    }

    private static async Task<string> CreateOneOnOneChatAsync(HttpClient client, string meId, string targetUserId)
    {
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

        using var body = new StringContent(payload.ToJsonString(), Encoding.UTF8, "application/json");
        using var response = await client.PostAsync($"{GraphBaseUrl}/chats", body);
        var content = await response.Content.ReadAsStringAsync();
        EnsureSuccess(response, content);

        using var doc = JsonDocument.Parse(content);
        var id = GetString(doc.RootElement, "id");
        return id ?? throw new InvalidOperationException("Graph did not return chat id.");
    }

    private static async Task<(List<JsonElement> Chats, int ScannedCount, int ScannedPageCount)> ListChatsAsync(
        HttpClient client,
        int? maxChats,
        string chatType,
        DateTimeOffset? startUtc,
        DateTimeOffset? endUtc,
        string topic,
        bool? isRead,
        int? maxPageCount)
    {
        var query = new Dictionary<string, string>
        {
            ["$top"] = "50",
            ["$orderby"] = "lastMessagePreview/createdDateTime desc",
            ["$select"] = "id,chatType,topic,lastUpdatedDateTime,lastMessagePreview,webUrl,viewpoint",
            ["$expand"] = "lastMessagePreview"
        };

        var chatFilter = BuildChatFilter(chatType, startUtc, endUtc, topic);
        if (!string.IsNullOrWhiteSpace(chatFilter))
        {
            query["$filter"] = chatFilter;
        }

        var url = BuildUrl($"{GraphBaseUrl}/me/chats", query);
        var chats = new List<JsonElement>();
        var scannedCount = 0;
        var scannedPageCount = 0;
        var unlimited = !maxChats.HasValue;
        var emptyPageBudget = -1;

        while (!string.IsNullOrWhiteSpace(url) && (unlimited || chats.Count < maxChats.Value))
        {
            if (maxPageCount.HasValue && scannedPageCount >= maxPageCount.Value)
            {
                break;
            }

            using var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();
            EnsureSuccess(response, content);
            scannedPageCount++;

            using var doc = JsonDocument.Parse(content);
            var root = doc.RootElement;
            var pageItems = root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array
                ? valueElement.EnumerateArray().Select(x => x.Clone()).ToList()
                : new List<JsonElement>();

            foreach (var item in pageItems)
            {
                scannedCount++;

                if (!MatchesChatReadFilter(item, isRead))
                {
                    continue;
                }

                chats.Add(item);
                if (!unlimited && chats.Count >= maxChats.Value)
                {
                    break;
                }
            }

            if (!unlimited && chats.Count >= maxChats.Value)
            {
                break;
            }

            if (emptyPageBudget == -1 && pageItems.Count == 0)
            {
                emptyPageBudget = 5;
            }
            else if (emptyPageBudget > -1)
            {
                emptyPageBudget--;
                if (emptyPageBudget <= 0)
                {
                    break;
                }
            }

            url = root.TryGetProperty("@odata.nextLink", out var nextLink) ? nextLink.GetString() : null;
        }

        return (chats, scannedCount, scannedPageCount);
    }

    private static async Task<List<JsonElement>> ListRecentChatMessagesAsync(HttpClient client, string chatId, int maxMessages, DateTimeOffset? startUtc, DateTimeOffset? endUtc)
    {
        var query = new Dictionary<string, string>
        {
            ["$top"] = "50",
            ["$orderby"] = "lastModifiedDateTime desc"
        };

        var messageTimeFilter = BuildDateTimeFilter("lastModifiedDateTime", startUtc, endUtc);
        if (!string.IsNullOrWhiteSpace(messageTimeFilter))
        {
            query["$filter"] = messageTimeFilter;
        }

        var url = BuildUrl($"{GraphBaseUrl}/me/chats/{chatId}/messages", query);
        var messages = new List<JsonElement>();

        while (!string.IsNullOrWhiteSpace(url) && messages.Count < maxMessages)
        {
            using var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();
            EnsureSuccess(response, content);

            using var doc = JsonDocument.Parse(content);
            var root = doc.RootElement;
            var pageItems = root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array
                ? valueElement.EnumerateArray().Select(x => x.Clone()).ToList()
                : new List<JsonElement>();

            foreach (var item in pageItems)
            {
                messages.Add(item);
                if (messages.Count >= maxMessages)
                {
                    break;
                }
            }

            url = root.TryGetProperty("@odata.nextLink", out var nextLink) ? nextLink.GetString() : null;
        }

        return messages;
    }

    private static JsonObject FormatMessage(JsonElement chat, JsonElement message, JsonObject me)
    {
        try
        {
            var createdText = GetString(message, "createdDateTime");
            var createdUtc = ParseGraphDateTime(createdText);
            var lastReadText = GetChatLastReadDateTime(chat);
            var lastReadUtc = ParseGraphDateTime(lastReadText);
            var myMention = ExtractMyMention(message, me);

            var sender = message.TryGetProperty("from", out var fromElement) &&
                         fromElement.ValueKind == JsonValueKind.Object &&
                         fromElement.TryGetProperty("user", out var userElement) &&
                         userElement.ValueKind == JsonValueKind.Object
                ? userElement
                : default;

            var chatObj = new JsonObject
            {
                ["id"] = GetString(chat, "id"),
                ["type"] = GetString(chat, "chatType"),
                ["topic"] = GetString(chat, "topic"),
                ["webUrl"] = GetString(chat, "webUrl"),
                ["lastMessageReadDateTime"] = lastReadText
            };

            return new JsonObject
            {
                ["chat"] = chatObj,
                ["id"] = GetString(message, "id"),
                ["createdDateTime"] = createdText,
                ["lastModifiedDateTime"] = GetString(message, "lastModifiedDateTime"),
                ["messageType"] = GetString(message, "messageType"),
                ["subject"] = GetString(message, "subject"),
                ["summary"] = GetString(message, "summary"),
                ["webUrl"] = GetString(message, "webUrl"),
                ["from"] = new JsonObject
                {
                    ["id"] = GetString(sender, "id"),
                    ["displayName"] = GetString(sender, "displayName"),
                    ["userIdentityType"] = GetString(sender, "userIdentityType")
                },
                ["body"] = message.TryGetProperty("body", out var bodyElement)
                    ? new JsonObject
                    {
                        ["contentType"] = GetString(bodyElement, "contentType"),
                        ["content"] = GetString(bodyElement, "content")
                    }
                    : null,
                ["mentions"] = myMention,
                ["readState"] = InferReadState(createdUtc, lastReadUtc)
            };
        }
        catch (Exception ex)
        {
            SystemEventLogger.Error($"FormatMessage failed: {ex.Message}");
            return new JsonObject();
        }
    }

    private static bool MatchesReadMessageFilter(JsonElement message, string targetUserId)
    {
        if (string.IsNullOrWhiteSpace(targetUserId))
        {
            return true;
        }

        if (!message.TryGetProperty("from", out var fromElement) ||
            fromElement.ValueKind != JsonValueKind.Object ||
            !fromElement.TryGetProperty("user", out var userElement) ||
            userElement.ValueKind != JsonValueKind.Object)
        {
            return false;
        }

        var senderId = GetString(userElement, "id");
        return string.Equals(senderId, targetUserId, StringComparison.OrdinalIgnoreCase);
    }

    private static JsonObject ExtractMyMention(JsonElement message, JsonObject me)
    {
        if (!message.TryGetProperty("mentions", out var mentions) || mentions.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        var myId = GetString(me, "id")?.Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(myId))
        {
            return null;
        }

        foreach (var mention in mentions.EnumerateArray())
        {
            if (!mention.TryGetProperty("mentioned", out var mentioned) ||
                mentioned.ValueKind != JsonValueKind.Object ||
                !mentioned.TryGetProperty("user", out var user) ||
                user.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var uid = GetString(user, "id")?.Trim().ToLowerInvariant();
            if (!string.IsNullOrWhiteSpace(uid) && uid == myId)
            {
                return new JsonObject
                {
                    ["displayName"] = GetString(user, "displayName"),
                    ["id"] = uid
                };
            }
        }

        return null;
    }

    private static string InferReadState(DateTimeOffset? messageCreatedUtc, DateTimeOffset? lastReadUtc)
    {
        if (messageCreatedUtc is null)
        {
            return "unknown";
        }

        if (lastReadUtc is null)
        {
            return "unread";
        }

        return messageCreatedUtc <= lastReadUtc ? "read" : "unread";
    }

    private static bool? IsChatRead(JsonElement chat)
    {
        var previewCreatedUtc = ParseGraphDateTime(GetChatLastMessagePreviewCreatedDateTime(chat));
        if (!previewCreatedUtc.HasValue)
        {
            return null;
        }

        var lastReadUtc = ParseGraphDateTime(GetChatLastReadDateTime(chat));
        if (!lastReadUtc.HasValue)
        {
            return false;
        }

        return previewCreatedUtc <= lastReadUtc;
    }

    private static bool MatchesChatReadFilter(JsonElement chat, bool? isRead)
    {
        if (!isRead.HasValue)
        {
            return true;
        }

        var currentChatRead = IsChatRead(chat);
        return currentChatRead.HasValue && currentChatRead.Value == isRead.Value;
    }

    private static string GetChatLastReadDateTime(JsonElement chat)
    {
        var top = GetString(chat, "lastMessageReadDateTime");
        if (!string.IsNullOrWhiteSpace(top))
        {
            return top;
        }

        if (!chat.TryGetProperty("viewpoint", out var viewpoint) || viewpoint.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        return GetString(viewpoint, "lastMessageReadDateTime");
    }

    private static string GetChatLastMessagePreviewCreatedDateTime(JsonElement chat)
    {
        if (!chat.TryGetProperty("lastMessagePreview", out var lastMessagePreview) ||
            lastMessagePreview.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        return GetString(lastMessagePreview, "createdDateTime");
    }

    private static DateTimeOffset? ParseGraphDateTime(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var cleaned = value.Replace("Z", "+00:00", StringComparison.OrdinalIgnoreCase);
        if (!DateTimeOffset.TryParse(cleaned, out var dt))
        {
            return null;
        }

        return dt.ToUniversalTime();
    }

    private static string BuildUrl(string baseUrl, Dictionary<string, string> query)
    {
        var qs = string.Join("&", query.Select(kvp =>
            $"{Uri.EscapeDataString(kvp.Key)}={Uri.EscapeDataString(kvp.Value)}"));
        return string.IsNullOrWhiteSpace(qs) ? baseUrl : $"{baseUrl}?{qs}";
    }

    private static HttpClient CreateClient(string token, bool preferTextBody = false)
    {
        var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (preferTextBody)
        {
            client.DefaultRequestHeaders.Add("Prefer", "outlook.body-content-type=\"text\"");
        }
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        return client;
    }

    private static string BuildDateTimeFilter(string fieldName, DateTimeOffset? startUtc, DateTimeOffset? endUtc)
    {
        var filters = new List<string>();

        if (startUtc.HasValue)
        {
            filters.Add($"{fieldName} gt {startUtc.Value:yyyy-MM-ddTHH:mm:ssZ}");
        }

        if (endUtc.HasValue)
        {
            filters.Add($"{fieldName} lt {endUtc.Value:yyyy-MM-ddTHH:mm:ssZ}");
        }

        return string.Join(" and ", filters);
    }

    private static string BuildChatFilter(string chatType, DateTimeOffset? startUtc, DateTimeOffset? endUtc, string topic)
    {
        var filters = new List<string>();
        var dateTimeFilter = BuildDateTimeFilter("lastUpdatedDateTime", startUtc, endUtc);
        if (!string.IsNullOrWhiteSpace(dateTimeFilter))
        {
            filters.Add(dateTimeFilter);
        }

        if (string.Equals(chatType, "group", StringComparison.OrdinalIgnoreCase))
        {
            filters.Add("(chatType eq 'group' or chatType eq 'meeting')");
        }
        else if (!string.IsNullOrWhiteSpace(chatType) && !string.Equals(chatType, "all", StringComparison.OrdinalIgnoreCase))
        {
            filters.Add($"chatType eq '{chatType}'");
        }

        if (string.Equals(chatType, "group", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(topic))
        {
            var escapedTopic = topic.Replace("'", "''", StringComparison.Ordinal);
            filters.Add($"contains(topic,'{escapedTopic}')");
        }

        return string.Join(" and ", filters);
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

    private static string GetString(JsonObject obj, string key)
    {
        if (!obj.TryGetPropertyValue(key, out var node) || node is null)
        {
            return null;
        }

        return node.GetValue<string>();
    }

    private static void EnsureSuccess(HttpResponseMessage response, string content)
    {
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Graph API error {(int)response.StatusCode}: {content}");
        }
    }
}
