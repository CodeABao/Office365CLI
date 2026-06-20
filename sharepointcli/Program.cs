namespace sharepointcli;

internal partial class Program
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
            SystemEventLogger.Info($"Command started: {command}");
            var auth = new GraphAuthenticator();

            switch (command)
            {
                case "read":
                    return await HandleReadAsync(auth, options);
                case "search":
                    return await HandleSearchAsync(auth, options);
                case "upload":
                    return await HandleUploadAsync(auth, options);
                case "list-drives":
                    return await HandleListDrivesAsync(auth, options);
                case "list-drive-items":
                    return await HandleListDriveItemsAsync(auth, options);
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

    public static async Task<int> HandleReadAsync(GraphAuthenticator auth, string[] args)
    {
        var options = ParseOptionMap(args);
        var sharepointUrl = GetRequired(options, "--file-url");
        var download = options.ContainsKey("--download");
        var downloadDirectory = GetOptional(options, "--download-dir");

        if (download && string.IsNullOrWhiteSpace(downloadDirectory))
        {
            throw new ArgumentException("--download-dir is required when --download is specified");
        }

        if (!download && !string.IsNullOrWhiteSpace(downloadDirectory))
        {
            throw new ArgumentException("--download-dir can only be used with --download");
        }

        SystemEventLogger.Info($"Reading SharePoint file. download={download}, hasDownloadDirectory={!string.IsNullOrWhiteSpace(downloadDirectory)}");

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var reader = new FileReader();
        return await reader.ReadAsync(token, sharepointUrl, download, downloadDirectory);
    }

    public static async Task<int> HandleSearchAsync(GraphAuthenticator auth, string[] args)
    {
        var options = ParseOptionMap(args);
        var keyword = GetRequired(options, "--keyword");
        var topText = GetOptional(options, "--top") ?? "10";

        if (!int.TryParse(topText, out var top) || top <= 0)
        {
            throw new ArgumentException("--top must be an integer > 0");
        }

        SystemEventLogger.Info($"Searching SharePoint files. keyword={keyword}, top={top}");

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var reader = new FileReader();
        return await reader.SearchAsync(token, keyword, top);
    }

    public static async Task<int> HandleUploadAsync(GraphAuthenticator auth, string[] args)
    {
        var options = ParseOptionMap(args);

        var siteUrl = GetOptional(options, "--site-url");
        var libraryName = GetOptional(options, "--library-name");
        var libraryUrl = GetOptional(options, "--library-url");
        var driveId = GetOptional(options, "--drive-id");
        var filePath = GetOptional(options, "--file-path");
        var text = GetOptional(options, "--text");
        var fileName = GetOptional(options, "--file-name");

        var libraryTargetCount = Convert.ToInt32(!string.IsNullOrWhiteSpace(libraryName))
            + Convert.ToInt32(!string.IsNullOrWhiteSpace(libraryUrl))
            + Convert.ToInt32(!string.IsNullOrWhiteSpace(driveId));
        if (libraryTargetCount != 1)
        {
            throw new ArgumentException("Specify exactly one of --library-name, --library-url, or --drive-id");
        }

        if (string.IsNullOrWhiteSpace(driveId) && string.IsNullOrWhiteSpace(siteUrl))
        {
            throw new ArgumentException("--site-url is required when using --library-name or --library-url");
        }

        var contentSourceCount = Convert.ToInt32(!string.IsNullOrWhiteSpace(filePath)) + Convert.ToInt32(!string.IsNullOrWhiteSpace(text));
        if (contentSourceCount != 1)
        {
            throw new ArgumentException("Specify exactly one of --file-path or --text");
        }

        if (!string.IsNullOrWhiteSpace(text) && string.IsNullOrWhiteSpace(fileName))
        {
            throw new ArgumentException("--file-name is required when using --text");
        }

        SystemEventLogger.Info($"Uploading SharePoint file. siteUrl={siteUrl}, libraryName={libraryName}, libraryUrl={libraryUrl}, driveId={driveId}, hasFilePath={!string.IsNullOrWhiteSpace(filePath)}, hasText={!string.IsNullOrWhiteSpace(text)}");

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var uploader = new FileUploader();
        return await uploader.UploadAsync(token, siteUrl, libraryName, libraryUrl, driveId, filePath, text, fileName);
    }

    public static async Task<int> HandleListDrivesAsync(GraphAuthenticator auth, string[] args)
    {
        var options = ParseOptionMap(args);
        var siteUrl = GetRequired(options, "--site-url");

        SystemEventLogger.Info($"Listing SharePoint drives. siteUrl={siteUrl}");

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var lister = new SiteDriveLister();
        return await lister.ListAsync(token, siteUrl);
    }

    public static async Task<int> HandleListDriveItemsAsync(GraphAuthenticator auth, string[] args)
    {
        var options = ParseOptionMap(args);
        var driveId = GetRequired(options, "--drive-id");
        var topText = GetOptional(options, "--top") ?? "20";
        var orderBy = GetOptional(options, "--order-by") ?? "name";
        var objectType = GetOptional(options, "--objectType") ?? "all";
        var desc = options.ContainsKey("--desc");

        if (!int.TryParse(topText, out var top) || top <= 0)
        {
            throw new ArgumentException("--top must be an integer > 0");
        }

        var allowedObjectTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "all",
            "folder",
            "file"
        };

        if (!allowedObjectTypes.Contains(objectType))
        {
            throw new ArgumentException("--objectType must be one of: all, folder, file");
        }

        var allowedOrderBy = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "name",
            "size",
            "createdDateTime",
            "lastModifiedDateTime"
        };

        if (!allowedOrderBy.Contains(orderBy))
        {
            throw new ArgumentException("--order-by must be one of: name, size, createdDateTime, lastModifiedDateTime");
        }

        SystemEventLogger.Info($"Listing drive items. driveId={driveId}, top={top}, orderBy={orderBy}, desc={desc}, objectType={objectType}");

        var token = await auth.AcquireAccessTokenAsync(["https://graph.microsoft.com/.default"]);
        var lister = new DriveItemLister();
        return await lister.ListAsync(token, driveId, objectType, top, orderBy, desc);
    }

    // 参数解析工具，与 outlookcli 风格一致
    private static Dictionary<string, string> ParseOptionMap(string[] args)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < args.Length; i++)
        {
            var key = args[i];
            if (!key.StartsWith("--", StringComparison.Ordinal))
                throw new ArgumentException($"Invalid argument: {key}");
            if (key == "--download" || key == "--desc")
            {
                map[key] = "true";
                continue;
            }
            if (i + 1 >= args.Length || args[i + 1].StartsWith("--", StringComparison.Ordinal))
                throw new ArgumentException($"Missing value for argument: {key}");

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
            return value;
        throw new ArgumentException($"Missing required argument: {key}");
    }
    private static string GetOptional(Dictionary<string, string> options, string key)
    {
        return options.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : null;
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:\n  sharepointcli read --file-url <url> [--download --download-dir <directory>]\n  sharepointcli search --keyword <text> [--top 10]\n  sharepointcli upload [--site-url <url>] (--library-name <name> | --library-url <url> | --drive-id <id>) (--file-path <path> | --text <text>) [--file-name <name>]\n  sharepointcli list-drives --site-url <url>\n  sharepointcli list-drive-items --drive-id <id> [--objectType all|folder|file] [--top 20] [--order-by name|size|createdDateTime|lastModifiedDateTime] [--desc]");
    }
}
