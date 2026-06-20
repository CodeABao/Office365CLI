namespace sharepointcli;

internal static class SystemEventLogger
{
    private static readonly string LogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Office365CLI",
        "sharepointcli",
        "Logs");

    public static void Info(string message)
    {
        Log("INFORMATION", message);
    }

    public static void Error(string message)
    {
        Log("ERROR", message);
    }

    private static void Log(string level, string message)
    {
        try
        {
            Directory.CreateDirectory(LogDirectory);

            var logFilePath = Path.Combine(
                LogDirectory,
                $"sharepointcli_{DateTime.UtcNow:yyyyMMdd}.log");

            File.AppendAllText(logFilePath, $"{DateTimeOffset.UtcNow:O} [{level}] {message}{Environment.NewLine}");
        }
        catch
        {
            // Best-effort logging only; never fail the main flow.
        }
    }
}