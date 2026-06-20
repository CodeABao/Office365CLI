using System.Globalization;
using System.Net.Mail;
using System.Text;

namespace outlookcli;

internal class Program
{
    public static async Task<int> Main(string[] args)
    {
        if (args.Length == 0)
        {
            PrintUsage();
            return 1;
        }

        SystemEventLogger.Info(string.Join(" ", args));

        var command = args[0].Trim().ToLowerInvariant();
        var options = args.Skip(1).ToArray();

        try
        {
            var auth = new GraphAuthenticator();

            switch (command)
            {
                case "send":
                    await HandleSendAsync(auth, options);
                    return 0;
                case "draft":
                    await HandleDraftAsync(auth, options);
                    return 0;
                case "list":
                    await HandleListAsync(auth, options);
                    return 0;
                case "read":
                    await HandleReadByIdAsync(auth, options);
                    return 0;
                case "reply":
                    await HandleReplyAsync(auth, options);
                    return 0;
                case "search":
                    await HandleSearchAsync(auth, options);
                    return 0;
                case "-h":
                case "--help":
                case "help":
                    PrintUsage();
                    return 0;
                case "--version":
                    Console.WriteLine("0.0.1");
                    return 0;
                default:
                    Console.Error.WriteLine($"Unknown command: {command}");
                    PrintUsage();
                    return 1;
            }
        }
        catch (Exception ex)
        {
            SystemEventLogger.Error($"Command failed: {command}. {ex.Message}");
            Console.Error.WriteLine($"ERROR: {ex.Message}");
            return 1;
        }
    }

    private static async Task HandleSendAsync(GraphAuthenticator auth, string[] args)
    {
        var map = ParseOptionMap(args);
        var to = NormalizeRecipientAddress(GetRequiredOption(map, "--to"));
        var subject = GetRequiredOption(map, "--subject");
        var body = GetRequiredOption(map, "--body");
        var bodyType = GetBodyContentType(map);

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var sender = new MailSender();
        await sender.SendAsync(token, to, subject, body, bodyType);
        Console.WriteLine("Mail sent successfully.");
    }

    private static async Task HandleDraftAsync(GraphAuthenticator auth, string[] args)
    {
        var map = ParseOptionMap(args);
        var messageId = GetOptionalTrimmedOption(map, "--id");
        var body = GetRequiredOption(map, "--body");
        var bodyType = GetBodyContentType(map);

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var draftCreator = new MailDraftCreator();

        string output;
        if (!string.IsNullOrWhiteSpace(messageId))
        {
            if (map.ContainsKey("--to") || map.ContainsKey("--subject"))
            {
                throw new ArgumentException("--to and --subject are not supported when --id is provided.");
            }

            output = await draftCreator.CreateReplyDraftAsync(token, messageId, body, bodyType);
        }
        else
        {
            var to = NormalizeRecipientAddress(GetRequiredOption(map, "--to"));
            var subject = GetRequiredOption(map, "--subject");
            output = await draftCreator.CreateDraftAsync(token, to, subject, body, bodyType);
        }

        Console.WriteLine(output);
    }

    private static async Task HandleListAsync(GraphAuthenticator auth, string[] args)
    {
        var map = ParseOptionMap(args);
        var start = GetOptionalDateTimeOption(map, "--start") ?? DateTimeOffset.UtcNow.AddMonths(-1);
        var end = GetOptionalDateTimeOption(map, "--end");
        var readStatus = GetOptionOrDefault(map, "--read-status", "all").ToLowerInvariant();
        var topText = GetOptionOrDefault(map, "--top", "50");
        var includeBody = GetBooleanOption(map, "--include-body", false);
        var folder = GetOptionalTrimmedOption(map, "--folder");

        if (end.HasValue && end < start)
        {
            throw new ArgumentException("--end must be greater than or equal to --start.");
        }

        if (readStatus is not ("all" or "read" or "unread"))
        {
            throw new ArgumentException("--read-status must be one of: all, read, unread.");
        }

        if (!int.TryParse(topText, out var top) || top is < 1 or > 1000)
        {
            throw new ArgumentException("--top must be an integer between 1 and 1000.");
        }

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var reader = new MailReader();
        var output = await reader.ListAsync(token, start, end, readStatus, top, includeBody, folder);
        Console.WriteLine(output);
    }

    private static async Task HandleSearchAsync(GraphAuthenticator auth, string[] args)
    {
        var map = ParseOptionMap(args);
        var start = GetOptionalDateTimeOption(map, "--start") ?? DateTimeOffset.UtcNow.AddMonths(-1);
        var end = GetOptionalDateTimeOption(map, "--end");
        var readStatus = GetOptionOrDefault(map, "--read-status", "all").ToLowerInvariant();
        var topText = GetOptionOrDefault(map, "--top", "50");
        var includeBody = GetBooleanOption(map, "--include-body", false);
        var sender = GetOptionalTrimmedOption(map, "--from");
        var subject = GetOptionalTrimmedOption(map, "--subject");

        if (end.HasValue && end < start)
        {
            throw new ArgumentException("--end must be greater than or equal to --start.");
        }

        if (readStatus is not ("all" or "read" or "unread"))
        {
            throw new ArgumentException("--read-status must be one of: all, read, unread.");
        }

        if (!int.TryParse(topText, out var top) || top is < 1 or > 1000)
        {
            throw new ArgumentException("--top must be an integer between 1 and 1000.");
        }

        if (string.IsNullOrWhiteSpace(sender) && string.IsNullOrWhiteSpace(subject))
        {
            throw new ArgumentException("search requires at least one of: --from, --subject.");
        }

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var reader = new MailReader();
        var output = await reader.SearchAsync(token, start, end, readStatus, top, sender, subject, includeBody);
        Console.WriteLine(output);
    }

    private static async Task HandleReadByIdAsync(GraphAuthenticator auth, string[] args)
    {
        var map = ParseOptionMap(args);
        var id = GetRequiredOption(map, "--id").Trim();

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var reader = new MailReader();
        var output = await reader.ReadByIdAsync(token, id);
        Console.WriteLine(output);
    }

    private static async Task HandleReplyAsync(GraphAuthenticator auth, string[] args)
    {
        var map = ParseOptionMap(args);
        var id = GetRequiredOption(map, "--id").Trim();
        var body = GetRequiredOption(map, "--body");
        var bodyType = GetBodyContentType(map);

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var replier = new MailReplier();
        await replier.ReplyAsync(token, id, body, bodyType);
        Console.WriteLine("Mail replied successfully.");
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

    private static string GetRequiredOption(Dictionary<string, string> map, string key)
    {
        if (map.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        throw new ArgumentException($"Missing required argument: {key}");
    }

    private static string GetOptionOrDefault(Dictionary<string, string> map, string key, string defaultValue)
    {
        return map.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : defaultValue;
    }

    private static bool GetBooleanOption(Dictionary<string, string> map, string key, bool defaultValue)
    {
        if (!map.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
        {
            return defaultValue;
        }

        if (bool.TryParse(value.Trim(), out var parsed))
        {
            return parsed;
        }

        throw new ArgumentException($"{key} must be either true or false.");
    }

    private static string GetOptionalTrimmedOption(Dictionary<string, string> map, string key)
    {
        return map.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value.Trim()
            : null;
    }

    private static DateTimeOffset? GetOptionalDateTimeOption(Dictionary<string, string> map, string key)
    {
        return map.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
            ? ParseDateTimeAsLocalOrOffset(value, key)
            : null;
    }

    private static DateTimeOffset ParseDateTimeAsLocalOrOffset(string value, string argumentName)
    {
        var trimmed = value.Trim();

        if (!DateTimeOffset.TryParse(
                trimmed,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal,
                out var parsed))
        {
            throw new ArgumentException($"Invalid datetime format for {argumentName}: {value}");
        }

        return parsed.ToUniversalTime();
    }

    private static string GetBodyContentType(Dictionary<string, string> map)
    {
        var bodyType = GetOptionOrDefault(map, "--body-type", "text").Trim().ToLowerInvariant();

        return bodyType switch
        {
            "text" => "Text",
            "html" => "HTML",
            _ => throw new ArgumentException("--body-type must be one of: text, html.")
        };
    }

    private static string NormalizeRecipientAddress(string recipient)
    {
        var trimmedRecipient = recipient.Trim();

        if (MailAddress.TryCreate(trimmedRecipient, out var parsedRecipient))
        {
            return parsedRecipient.Address;
        }

        throw new ArgumentException($"Invalid email address format for --to: {recipient}. Expected a complete email address like user@example.com.");
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:\n  outlookcli send --to user@example.com --subject \"Test\" --body \"Hello!\" [--body-type text]\n  outlookcli draft --to user@example.com --subject \"Test\" --body \"Hello!\" [--body-type text]\n  outlookcli draft --id AAMkAG... --body \"Thanks\" [--body-type text]\n  outlookcli list [--start 2026-05-01 00:00:00] [--end 2026-05-13 23:59:59] [--read-status all] [--top 100] [--include-body false] [--folder inbox]\n  outlookcli read --id AAMkAG...\n  outlookcli reply --id AAMkAG... --body \"Thanks\" [--body-type text]\n  outlookcli search [--from sender@example.com] [--subject \"keyword\"] [--start 2026-05-01 00:00:00] [--end 2026-05-13 23:59:59] [--read-status all] [--top 100] [--include-body false]\n\nWhen --to is not a complete email address, an error is returned.");
    }
}
