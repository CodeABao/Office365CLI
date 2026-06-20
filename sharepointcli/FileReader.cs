using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace sharepointcli;

using System.IO;
using System.Text;
using System.Text.Json;
using System.Net.Http.Headers;
using System.IO.Packaging;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml.Packaging; // 需安装DocumentFormat.OpenXml
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

internal sealed class FileReader
{
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    public async Task<int> SearchAsync(string token, string keyword, int top)
    {
        using var client = CreateClient(token);

        var payload = new JsonObject
        {
            ["requests"] = new JsonArray
            {
                new JsonObject
                {
                    ["entityTypes"] = new JsonArray("driveItem"),
                    ["query"] = new JsonObject
                    {
                        ["queryString"] = keyword
                    },
                    ["from"] = 0,
                    ["size"] = top
                }
            }
        };

        var json = payload.ToJsonString();
        using var body = new StringContent(json, Encoding.UTF8, "application/json");
        using var response = await client.PostAsync($"{GraphBaseUrl}/search/query", body);
        var content = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Failed to search SharePoint files. Graph API error {(int)response.StatusCode}: {content}");
        }

        using var doc = JsonDocument.Parse(content);
        var hits = ExtractSearchHits(doc.RootElement);

        var output = new JsonObject
        {
            ["count"] = hits.Count,
            ["query"] = new JsonObject
            {
                ["keyword"] = keyword,
                ["top"] = top
            },
            ["results"] = hits
        };

        Console.WriteLine(output.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        return 0;
    }

    public async Task<int> ReadAsync(string token, string sharePointUrl, bool download, string downloadDirectory)
    {
        var item = await GetDriveItemAsync(token, sharePointUrl);
        var content = await GetFileContentAsync(token, sharePointUrl);

        var name = GetString(item, "name") ?? string.Empty;
        var size = GetLong(item, "size");
        var webUrl = GetString(item, "webUrl") ?? string.Empty;
        var createdBy = GetCreatedByDisplayName(item);

        Console.WriteLine($"Resolved file: {name}");
        Console.WriteLine($"Size: {size} bytes");
        if (!string.IsNullOrWhiteSpace(createdBy))
        {
            Console.WriteLine($"Created by: {createdBy}");
        }
        if (!string.IsNullOrWhiteSpace(webUrl))
        {
            Console.WriteLine($"Web URL: {webUrl}");
        }

        SystemEventLogger.Info($"Resolved SharePoint file: {name}, size={size}, download={download}");

        if (download)
        {
            await SaveToDirectoryAsync(content, downloadDirectory, name);
            return 0;
        }

        await PrintContentAsync(content, name);
        return 0;
    }

    private static string ToShareId(string sharePointUrl)
    {
        var bytes = Encoding.UTF8.GetBytes(sharePointUrl);
        var encoded = Convert.ToBase64String(bytes)
            .Replace('+', '-')
            .Replace('/', '_')
            .TrimEnd('=');

        return $"u!{encoded}";
    }

    private async Task<JsonElement> GetDriveItemAsync(string token, string sharePointUrl)
    {
        var shareId = ToShareId(sharePointUrl);
        var url = $"{GraphBaseUrl}/shares/{shareId}/driveItem";

        using var client = CreateClient(token);
        using var response = await client.GetAsync(url);
        var content = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Failed to resolve SharePoint URL. Graph API error {(int)response.StatusCode}: {content}");
        }

        using var doc = JsonDocument.Parse(content);
        return doc.RootElement.Clone();
    }

    private async Task<byte[]> GetFileContentAsync(string token, string sharePointUrl)
    {
        var shareId = ToShareId(sharePointUrl);
        var url = $"{GraphBaseUrl}/shares/{shareId}/driveItem/content";

        using var client = CreateClient(token);
        using var response = await client.GetAsync(url);

        if (!response.IsSuccessStatusCode)
        {
            var content = await response.Content.ReadAsStringAsync();
            throw new InvalidOperationException($"Failed to fetch file content. Graph API error {(int)response.StatusCode}: {content}");
        }

        return await response.Content.ReadAsByteArrayAsync();
    }

    private static async Task SaveToDirectoryAsync(byte[] content, string downloadDirectory, string fileName)
    {
        var targetDirectory = Path.GetFullPath(downloadDirectory ?? throw new ArgumentException("Missing download directory."));
        Directory.CreateDirectory(targetDirectory);
        var outputPath = Path.Combine(targetDirectory, fileName);
        await File.WriteAllBytesAsync(outputPath, content);
        Console.WriteLine($"Saved file content to: {outputPath}");
    }

    private static async Task PrintContentAsync(byte[] content, string fileName)
    {
        await Task.CompletedTask;

        var suffix = Path.GetExtension(fileName).ToLowerInvariant();
        string text = null;
        string warning = null;
        try
        {
            switch (suffix)
            {
                case ".txt":
                case ".md":
                case ".json":
                case ".csv":
                case ".log":
                case ".xml":
                case ".html":
                case ".htm":
                    text = Encoding.UTF8.GetString(content);
                    break;
                case ".docx":
                    text = ExtractDocxText(content);
                    break;
                case ".xlsx":
                    text = ExtractXlsxText(content);
                    break;
                case ".pptx":
                    text = ExtractPptxText(content);
                    break;
                case ".pdf":
                    text = ExtractPdfText(content);
                    break;
                default:
                    warning = $"无法读取内容: 不支持的文件类型 {suffix}";
                    break;
            }
        }
        catch (Exception ex)
        {
            SystemEventLogger.Error($"Content extraction failed for {fileName}: {ex.Message}");
            warning = $"内容解析失败: {ex.Message}";
        }

        if (!string.IsNullOrWhiteSpace(warning))
        {
            SystemEventLogger.Info($"Content output warning for {fileName}: {warning}");
            Console.Error.WriteLine(warning);
            Console.Error.WriteLine("提示: 请使用 --download --download-dir <directory> 下载原文件。");
            Console.WriteLine($"Bytes length: {content.Length}");
            return;
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            Console.WriteLine("No text extracted.");
            return;
        }

        Console.WriteLine(text);
    }


    // Office/PDF文本提取实现
    private static string ExtractDocxText(byte[] content)
    {
        using var ms = new MemoryStream(content);
        using var doc = WordprocessingDocument.Open(ms, false);
        var sb = new StringBuilder();
        foreach (var text in doc.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
            sb.Append(text.Text);
        return sb.ToString();
    }

    private static string ExtractXlsxText(byte[] content)
    {
        using var ms = new MemoryStream(content);
        using var doc = SpreadsheetDocument.Open(ms, false);
        var sb = new StringBuilder();
        foreach (var sheet in doc.WorkbookPart.WorksheetParts)
        {
            foreach (var row in sheet.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>())
            {
                foreach (var cell in row.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                {
                    sb.Append(cell.InnerText);
                    sb.Append('\t');
                }
                sb.AppendLine();
            }
        }
        return sb.ToString();
    }

    private static string ExtractPptxText(byte[] content)
    {
        using var ms = new MemoryStream(content);
        using var doc = PresentationDocument.Open(ms, false);
        var sb = new StringBuilder();
        foreach (var slide in doc.PresentationPart.SlideParts)
        {
            foreach (var text in slide.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                sb.AppendLine(text.Text);
        }
        return sb.ToString();
    }

    private static string ExtractPdfText(byte[] content)
    {
        using var ms = new MemoryStream(content);
        using var pdf = PdfDocument.Open(ms);
        var sb = new StringBuilder();
        foreach (Page page in pdf.GetPages())
        {
            sb.AppendLine(page.Text);
        }
        return sb.ToString();
    }

    private static HttpClient CreateClient(string token)
    {
        var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        return client;
    }

    private static string GetString(JsonElement root, string propertyName)
    {
        return root.TryGetProperty(propertyName, out var value) && value.ValueKind != JsonValueKind.Null
            ? value.GetString()
            : null;
    }

    private static long GetLong(JsonElement root, string propertyName)
    {
        return root.TryGetProperty(propertyName, out var value) && value.TryGetInt64(out var parsed)
            ? parsed
            : 0;
    }

    private static string GetCreatedByDisplayName(JsonElement root)
    {
        if (!root.TryGetProperty("createdBy", out var createdBy) || createdBy.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        if (createdBy.TryGetProperty("user", out var user) && user.ValueKind == JsonValueKind.Object)
        {
            var displayName = GetString(user, "displayName");
            if (!string.IsNullOrWhiteSpace(displayName))
            {
                return displayName;
            }
        }

        if (createdBy.TryGetProperty("application", out var application) && application.ValueKind == JsonValueKind.Object)
        {
            return GetString(application, "displayName");
        }

        return null;
    }

    private static JsonArray ExtractSearchHits(JsonElement root)
    {
        var results = new JsonArray();

        if (!root.TryGetProperty("value", out var valueElement) || valueElement.ValueKind != JsonValueKind.Array)
        {
            return results;
        }

        foreach (var response in valueElement.EnumerateArray())
        {
            if (!response.TryGetProperty("hitsContainers", out var containersElement) || containersElement.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            foreach (var container in containersElement.EnumerateArray())
            {
                if (!container.TryGetProperty("hits", out var hitsElement) || hitsElement.ValueKind != JsonValueKind.Array)
                {
                    continue;
                }

                foreach (var hit in hitsElement.EnumerateArray())
                {
                    if (!hit.TryGetProperty("resource", out var resource) || resource.ValueKind != JsonValueKind.Object)
                    {
                        continue;
                    }

                    results.Add(new JsonObject
                    {
                        ["id"] = GetString(resource, "id"),
                        ["name"] = GetString(resource, "name"),
                        ["createdBy"] = GetCreatedByDisplayName(resource),
                        ["webUrl"] = GetString(resource, "webUrl"),
                        ["size"] = GetLong(resource, "size"),
                        ["createdDateTime"] = GetString(resource, "createdDateTime"),
                        ["lastModifiedDateTime"] = GetString(resource, "lastModifiedDateTime"),
                        ["summary"] = hit.TryGetProperty("summary", out var summaryElement) && summaryElement.ValueKind != JsonValueKind.Null
                            ? summaryElement.GetString()
                            : null
                    });
                }
            }
        }

        return results;
    }
}
