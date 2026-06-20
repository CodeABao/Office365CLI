using System.Globalization;
using System.Net.Mail;

namespace teamscli;

internal class Program
{
    private const string WindowsShareExperienceProdClientId = "a8759234-4b8b-4d94-8c0a-ee1ab73af270";
    private const string OfficeClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c";
    private const string OutlookClientId = "5d661950-3475-41cd-a2c3-d671a3162bc1";

    //目前不可用
    private const string SalesInsightsWebAppClientId = "b20d0d3a-dc90-485b-ad11-6031e769e221";

    public static async Task<int> Main(string[] args)
    {
        if (args.Length == 0)
        {
            PrintUsage();
            return 1;
        }

        SystemEventLogger.Info(string.Join(" ", args));

        var command = args[0].ToLowerInvariant();
        var options = ParseOptionMap(args.Skip(1).ToArray());

        try
        {
            SystemEventLogger.Info($"Command started: {command}");
            var auth = new GraphAuthenticator();

            switch (command)
            {
                case "send-message":
                    await HandleSendMessageAsync(auth, options);
                    return 0;
                case "send-channel-message":
                    await HandleSendChannelMessageAsync(auth, options);
                    return 0;
                case "list-messages":
                    await HandleListMessagesAsync(auth, options);
                    return 0;
                case "list-chats":
                    await HandleListChatsAsync(auth, options);
                    return 0;
                case "list-events":
                    await HandleListCalendarAsync(auth, options);
                    return 0;
                case "read-event":
                    await HandleReadEventAsync(auth, options);
                    return 0;
                case "list-joined-teams":
                    await HandleListJoinedTeamsAsync(auth, options);
                    return 0;
                case "list-team-channels":
                    await HandleListTeamChannelsAsync(auth, options);
                    return 0;
                case "list-channel-messages":
                    await HandleListChannelMessagesAsync(auth, options);
                    return 0;
                case "read-meeting-transcript":
                    await HandleReadMeetingTranscriptAsync(auth, options);
                    return 0;
                case "help":
                case "--help":
                case "-h":
                    PrintUsage();
                    return 0;
                case "--version":
                    Console.WriteLine("0.0.1");
                    return 0;
                default:
                    SystemEventLogger.Error($"Unknown command: {command}");
                    throw new ArgumentException($"Unknown command: {command}");
            }
        }
        catch (Exception ex)
        {
            SystemEventLogger.Error($"Command failed: {command}. {ex.Message}");
            Console.Error.WriteLine($"ERROR: {ex.Message}");
            PrintUsage();
            return 1;
        }
    }

    private static async Task HandleSendMessageAsync(GraphAuthenticator auth, Dictionary<string, string> options)
    {
        var hasUpn = options.ContainsKey("--upn");
        var hasId = options.ContainsKey("--id");
        var hasChatId = options.ContainsKey("--chat-id");
        var targetCount = Convert.ToInt32(hasUpn) + Convert.ToInt32(hasId) + Convert.ToInt32(hasChatId);
        if (targetCount != 1)
        {
            throw new ArgumentException("Specify exactly one of --upn, --id, or --chat-id");
        }

        var message = GetRequired(options, "--message");
        var contentType = GetOrDefault(options, "--content-type", "html");

        var graphToken = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);

        var o365Auth = new GraphAuthenticator(WindowsShareExperienceProdClientId);
        var o365Token = await o365Auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);

        var sender = new TeamsMessageSender();
        if (hasUpn)
        {
            var upn = NormalizeUpn(GetRequired(options, "--upn"));
            await sender.SendByUpnAsync(graphToken, o365Token, upn, message, contentType);
            return;
        }

        if (hasChatId)
        {
            var chatId = GetRequired(options, "--chat-id");
            await sender.SendToChatAsync(graphToken, chatId, message, contentType);
            return;
        }

        var id = GetRequired(options, "--id");
        await sender.SendByIdAsync(graphToken, o365Token, id, message, contentType);
    }

    private static async Task HandleSendChannelMessageAsync(GraphAuthenticator auth, Dictionary<string, string> options)
    {
        var teamId = GetRequired(options, "--team-id");
        var channelId = GetRequired(options, "--channel-id");
        var message = GetRequired(options, "--message");
        var contentType = GetOrDefault(options, "--content-type", "html");

        var o365Auth = new GraphAuthenticator(OfficeClientId);
        var token = await o365Auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);

        var sender = new TeamsMessageSender();
        await sender.SendToChannelAsync(token, teamId, channelId, message, contentType);
    }

    private static async Task HandleListMessagesAsync(GraphAuthenticator auth, Dictionary<string, string> options)
    {
        var hasUpn = options.ContainsKey("--upn");
        var hasChatId = options.ContainsKey("--chat-id");

        if (hasUpn == hasChatId)
        {
            throw new ArgumentException("Specify exactly one of --upn or --chat-id");
        }

        var topText = GetOrDefault(options, "--top", "50");
        if (!int.TryParse(topText, out var top) || top <= 0 || top > 500)
        {
            throw new ArgumentException("--top must be between 1 and 500");
        }

        var isReadText = GetOrDefault(options, "--is-read", "all").ToLowerInvariant();
        if (isReadText is not ("all" or "true" or "false"))
        {
            throw new ArgumentException("--is-read must be one of: all, true, false");
        }

        bool? isRead = isReadText switch
        {
            "true" => true,
            "false" => false,
            _ => null
        };

        var start = GetOptionalLocalDateTimeAsUtc(options, "--start");
        var end = GetOptionalLocalDateTimeAsUtc(options, "--end");

        if (start.HasValue && end.HasValue && end < start)
        {
            throw new ArgumentException("--end must be greater than or equal to --start");
        }

        var upn = hasUpn ? NormalizeUpn(GetRequired(options, "--upn")) : null;
        var chatId = hasChatId ? GetRequired(options, "--chat-id") : null;

        var o365Auth = new GraphAuthenticator(WindowsShareExperienceProdClientId);
        var token = await o365Auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var reader = new TeamsMessagesReader();
        var output = await reader.ListMessagesAsync(token, upn, chatId, isRead, top, start, end);
        Console.WriteLine(output);
    }

    private static async Task HandleListChatsAsync(GraphAuthenticator auth, Dictionary<string, string> options)
    {
        var name = GetOrDefault(options, "--name", string.Empty);
        var chatType = GetOrDefault(options, "--chat-type", "all");
        var topText = GetOrDefault(options, "--top", "50");
        if (!int.TryParse(topText, out var top) || top <= 0 || top > 500)
        {
            throw new ArgumentException("--top must be between 1 and 500");
        }

        if (chatType is not ("all" or "oneOnOne" or "group" or "meeting"))
        {
            throw new ArgumentException("--chat-type must be one of: all, oneOnOne, group, meeting");
        }

        var isReadText = GetOrDefault(options, "--is-read", "all").ToLowerInvariant();
        if (isReadText is not ("all" or "true" or "false"))
        {
            throw new ArgumentException("--is-read must be one of: all, true, false");
        }

        bool? isRead = isReadText switch
        {
            "true" => true,
            "false" => false,
            _ => null
        };

        var o365Auth = new GraphAuthenticator(WindowsShareExperienceProdClientId);
        var token = await o365Auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var reader = new TeamsMessagesReader();
        var output = await reader.ListChatsAsync(token, name, chatType, isRead, top);
        Console.WriteLine(output);
    }

    private static async Task HandleListCalendarAsync(GraphAuthenticator auth, Dictionary<string, string> options)
    {
        await HandleListCalendarEventsCoreAsync(auth, options);
    }

    private static async Task HandleReadEventAsync(GraphAuthenticator auth, Dictionary<string, string> options)
    {
        var eventId = GetRequired(options, "--event-id");
        var start = GetOptionalLocalDateTimeAsUtc(options, "--start");
        var end = GetOptionalLocalDateTimeAsUtc(options, "--end");

        if (start.HasValue && end.HasValue && end < start)
        {
            throw new ArgumentException("--end must be greater than or equal to --start");
        }

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var calendarReader = new TeamsCalendarReader();
        await calendarReader.ReadEventAsync(
            token,
            eventId,
            true,
            start?.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture),
            end?.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture));
    }

    private static async Task HandleListCalendarEventsCoreAsync(GraphAuthenticator auth, Dictionary<string, string> options)
    {
        var topText = GetOrDefault(options, "--top", "20");
        if (!int.TryParse(topText, out var top) || top <= 0)
        {
            throw new ArgumentException("--top must be greater than 0");
        }

        var includeBodyText = GetOrDefault(options, "--include-body", "false").ToLowerInvariant();
        if (includeBodyText is not ("true" or "false"))
        {
            throw new ArgumentException("--include-body must be true or false");
        }

        var includeBody = includeBodyText == "true";
        var subject = GetOrDefault(options, "--subject", string.Empty);
        var organizer = GetOrDefault(options, "--organizer", string.Empty);

        var start = GetOptionalLocalDateTimeAsUtc(options, "--start");
        var end = GetOptionalLocalDateTimeAsUtc(options, "--end");

        if (start.HasValue && end.HasValue && end < start)
        {
            throw new ArgumentException("--end must be greater than or equal to --start");
        }

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var calendarReader = new TeamsCalendarReader();
        await calendarReader.ReadAsync(
            token,
            top,
            includeBody,
            subject,
            organizer,
            start?.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture),
            end?.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture));
    }

    private static async Task HandleListJoinedTeamsAsync(GraphAuthenticator auth, Dictionary<string, string> options)
    {
        var name = GetOrDefault(options, "--name", string.Empty);
        var topText = GetOrDefault(options, "--top", "50");
        if (!int.TryParse(topText, out var top) || top <= 0 || top > 500)
        {
            throw new ArgumentException("--top must be between 1 and 500");
        }

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var reader = new TeamsMessagesReader();
        var output = await reader.ListJoinedTeamsAsync(token, name, top);
        Console.WriteLine(output);
    }

    private static async Task HandleListTeamChannelsAsync(GraphAuthenticator auth, Dictionary<string, string> options)
    {
        var teamId = GetRequired(options, "--team-id");
        var name = GetOrDefault(options, "--name", string.Empty);

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var reader = new TeamsMessagesReader();
        var output = await reader.ListTeamChannelsAsync(token, teamId, name);
        Console.WriteLine(output);
    }

    private static async Task HandleListChannelMessagesAsync(GraphAuthenticator auth, Dictionary<string, string> options)
    {
        var teamId = GetRequired(options, "--team-id");
        var channelId = GetRequired(options, "--channel-id");
        var topText = GetOrDefault(options, "--top", "50");
        if (!int.TryParse(topText, out var top) || top <= 0 || top > 500)
        {
            throw new ArgumentException("--top must be between 1 and 500");
        }

        var o365Auth = new GraphAuthenticator(OfficeClientId);
        var token = await o365Auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);

        var reader = new TeamsMessagesReader();
        var output = await reader.ListChannelMessagesAsync(token, teamId, channelId, top);
        Console.WriteLine(output);
    }

    private static async Task HandleReadMeetingTranscriptAsync(GraphAuthenticator auth, Dictionary<string, string> options)
    {
        var joinWebUrl = GetRequired(options, "--join-web-url");

        var outlookAuth = new GraphAuthenticator(OutlookClientId);
        var outlookToken = await outlookAuth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);

        var salesInsightsWebAppAuth = new GraphAuthenticator(SalesInsightsWebAppClientId);
        var salesInsightsWebAppToken = await salesInsightsWebAppAuth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);

        var reader = new MeetingTranscriptReader();
        var output = await reader.ReadTranscriptAsync(outlookToken, salesInsightsWebAppToken, joinWebUrl);
        Console.WriteLine(output);
    }

    private static Dictionary<string, string> ParseOptionMap(string[] args)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < args.Length; i++)
        {
            var key = args[i];
            if (!key.StartsWith("--", StringComparison.Ordinal))
            {
                throw new ArgumentException($"Invalid argument: {key}");
            }

            if (i + 1 >= args.Length || args[i + 1].StartsWith("--", StringComparison.Ordinal))
            {
                throw new ArgumentException($"Missing value for argument: {key}");
            }

            var valueParts = new List<string>();
            while (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
            {
                valueParts.Add(args[++i]);
            }

            map[key] = string.Join(" ", valueParts);
        }

        return map;
    }

    private static string GetRequired(Dictionary<string, string> options, string key)
    {
        if (options.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        throw new ArgumentException($"Missing required argument: {key}");
    }

    private static string GetOrDefault(Dictionary<string, string> options, string key, string defaultValue)
    {
        return options.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : defaultValue;
    }

    private static string NormalizeUpn(string upn)
    {
        var trimmedUpn = upn.Trim();

        if (MailAddress.TryCreate(trimmedUpn, out var parsedUpn))
        {
            return parsedUpn.Address;
        }

        throw new ArgumentException($"Invalid email address format for --upn: {upn}. Expected a complete email address like user@example.com.");
    }

    private static DateTimeOffset? GetOptionalLocalDateTimeAsUtc(Dictionary<string, string> options, string key)
    {
        return options.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
            ? ParseLocalDateTimeAsUtc(value, key)
            : null;
    }

    private static DateTimeOffset ParseLocalDateTimeAsUtc(string value, string argName)
    {
        var trimmed = value.Trim();

        if (!DateTimeOffset.TryParse(
                trimmed,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal,
                out var parsed))
        {
            throw new ArgumentException($"Invalid datetime format for {argName}: {value}. Expected local time like 2026-05-15 09:00:00");
        }

        return parsed.ToUniversalTime();
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:\n  teamscli send-message --upn user@example.com --message \"Hello\" [--content-type text]\n  teamscli send-message --id 00000000-0000-0000-0000-000000000000 --message \"Hello\" [--content-type text]\n  teamscli send-message --chat-id 19:examplethread@thread.v2 --message \"Hello group\" [--content-type text]\n  teamscli send-channel-message --team-id 00000000-0000-0000-0000-000000000000 --channel-id 19:channel@thread.tacv2 --message \"Hello channel\" [--content-type text]\n  teamscli list-messages --upn user@example.com [--is-read all|true|false] [--start \"2026-05-01 00:00:00\"] [--end \"2026-05-02 00:00:00\"] [--top 50]\n  teamscli list-messages --chat-id 19:examplethread@thread.v2 [--is-read all|true|false] [--start \"2026-05-01 00:00:00\"] [--end \"2026-05-02 00:00:00\"] [--top 50]\n  teamscli list-chats [--name \"Project Alpha\"] [--chat-type all|oneOnOne|group|meeting] [--is-read all|true|false] [--top 50]\n  teamscli list-events --top 20 [--include-body true|false] [--subject \"Quarterly Review\"] [--organizer \"alex@example.com\"] [--start \"2026-05-01 00:00:00\"] [--end \"2026-05-31 23:59:59\"]\n  teamscli read-event --event-id AAMkAG... [--start \"2026-05-01 00:00:00\"] [--end \"2026-05-31 23:59:59\"]\n  teamscli list-joined-teams [--name \"Engineering\"] [--top 50]\n  teamscli list-team-channels --team-id 00000000-0000-0000-0000-000000000000 [--name \"General\"]\n  teamscli list-channel-messages --team-id 00000000-0000-0000-0000-000000000000 --channel-id 19:channel@thread.tacv2 [--top 50]\n  teamscli read-meeting-transcript --join-web-url \"https://teams.microsoft.com/l/meetup-join/19%3ameeting_xxx%40thread.v2/0?context=...\"\n\nA complete email address (e.g. user@example.com) is required for --upn.");
    }
}
