using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using DocumentFormat.OpenXml.Packaging;

namespace Pactum.Showcase.Services;

public class DriveFileService
{
    private readonly DriveService _drive;
    private readonly string _rootFolderId;
    private readonly ILogger<DriveFileService> _logger;

    // Cache: city -> (folderId)
    private readonly Dictionary<string, string> _cityFolderIds = new(StringComparer.OrdinalIgnoreCase);
    // Cache: externalId -> folderId
    private readonly Dictionary<string, string> _bizFolderIds = new(StringComparer.OrdinalIgnoreCase);
    private readonly SemaphoreSlim _lock = new(1, 1);

    public DriveFileService(IConfiguration config, ILogger<DriveFileService> logger)
    {
        _logger = logger;
        _rootFolderId = config["GoogleDrive:RootFolderId"] ?? "1YOTXieCwwv6NXDJtmZC4jIYFXWql-wIu";

        var credPath = config["GoogleSheets:CredentialsPath"];
        var credJson = config["GoogleSheets:CredentialsJson"];

        GoogleCredential credential;
        if (!string.IsNullOrWhiteSpace(credJson))
            credential = GoogleCredential.FromJson(credJson);
        else if (!string.IsNullOrWhiteSpace(credPath) && File.Exists(credPath))
            credential = GoogleCredential.FromFile(credPath);
        else
            throw new InvalidOperationException("Google credentials not configured");

        credential = credential.CreateScoped(DriveService.Scope.DriveReadonly);
        _drive = new DriveService(new BaseClientService.Initializer { HttpClientInitializer = credential });
    }

    /// <summary>
    /// Find the business folder by externalId (e.g. "ID021539") in the city folder (e.g. "Москва").
    /// </summary>
    private async Task<string?> FindBizFolderAsync(string externalId, string city = "Москва")
    {
        if (string.IsNullOrWhiteSpace(externalId))
            return null;

        if (_bizFolderIds.TryGetValue(externalId, out var cached))
            return cached;

        await _lock.WaitAsync();
        try
        {
            if (_bizFolderIds.TryGetValue(externalId, out cached))
                return cached;

            var cityFolderId = await GetCityFolderIdAsync(city);
            if (cityFolderId == null)
                return null;

            // Search for folder starting with externalId
            var req = _drive.Files.List();
            req.Q = $"'{cityFolderId}' in parents and mimeType = 'application/vnd.google-apps.folder' and name contains '{externalId}'";
            req.Fields = "files(id, name)";
            var result = await req.ExecuteAsync();

            var folder = result.Files.FirstOrDefault(f => f.Name.StartsWith(externalId, StringComparison.OrdinalIgnoreCase));
            if (folder != null)
            {
                _bizFolderIds[externalId] = folder.Id;
                return folder.Id;
            }

            return null;
        }
        finally
        {
            _lock.Release();
        }
    }

    private async Task<string?> GetCityFolderIdAsync(string city)
    {
        if (_cityFolderIds.TryGetValue(city, out var cached))
            return cached;

        var req = _drive.Files.List();
        req.Q = $"'{_rootFolderId}' in parents and mimeType = 'application/vnd.google-apps.folder'";
        req.Fields = "files(id, name)";
        var result = await req.ExecuteAsync();

        foreach (var f in result.Files)
            _cityFolderIds[f.Name] = f.Id;

        return _cityFolderIds.GetValueOrDefault(city);
    }

    /// <summary>
    /// Read the main doc file (named same as folder) and return its text content as HTML paragraphs.
    /// </summary>
    public async Task<(string? html, string? error)> ReadMainDocAsync(string externalId)
    {
        var folderId = await FindBizFolderAsync(externalId);
        if (folderId == null)
            return (null, "Папка не найдена на Google Drive");

        // Find doc file with same name as folder (or starting with externalId)
        var req = _drive.Files.List();
        req.Q = $"'{folderId}' in parents and (mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' or mimeType = 'application/msword' or mimeType = 'application/vnd.google-apps.document')";
        req.Fields = "files(id, name, mimeType)";
        var files = await req.ExecuteAsync();

        // Find main doc — name starts with ID but does NOT contain "карточка"
        var mainDoc = files.Files.FirstOrDefault(f =>
            f.Name.StartsWith(externalId, StringComparison.OrdinalIgnoreCase)
            && !f.Name.Contains("карточка", StringComparison.OrdinalIgnoreCase));

        if (mainDoc == null)
            return (null, "Файл описания не найден");

        return await ReadDocContentAsync(mainDoc);
    }

    /// <summary>
    /// Read the card doc file (ID_карточка.docx).
    /// </summary>
    public async Task<(string? html, string? error)> ReadCardDocAsync(string externalId)
    {
        var folderId = await FindBizFolderAsync(externalId);
        if (folderId == null)
            return (null, "Папка не найдена на Google Drive");

        var req = _drive.Files.List();
        req.Q = $"'{folderId}' in parents and (mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' or mimeType = 'application/msword' or mimeType = 'application/vnd.google-apps.document')";
        req.Fields = "files(id, name, mimeType)";
        var files = await req.ExecuteAsync();

        var cardDoc = files.Files.FirstOrDefault(f =>
            f.Name.Contains("карточка", StringComparison.OrdinalIgnoreCase));

        if (cardDoc == null)
            return (null, "Файл карточки не найден");

        return await ReadDocContentAsync(cardDoc);
    }

    /// <summary>
    /// List image files in the business folder.
    /// </summary>
    public async Task<(List<ImageInfo> images, string? error)> ListImagesAsync(string externalId)
    {
        var folderId = await FindBizFolderAsync(externalId);
        if (folderId == null)
            return ([], "Папка не найдена на Google Drive");

        var req = _drive.Files.List();
        req.Q = $"'{folderId}' in parents and (mimeType contains 'image/')";
        req.Fields = "files(id, name, mimeType, thumbnailLink, webContentLink)";
        var result = await req.ExecuteAsync();

        var images = result.Files.Select(f => new ImageInfo
        {
            Id = f.Id,
            Name = f.Name,
            MimeType = f.MimeType,
            ThumbnailUrl = f.ThumbnailLink,
        }).ToList();

        return (images, null);
    }

    /// <summary>
    /// Get image bytes by file ID.
    /// </summary>
    public async Task<(byte[] data, string mimeType)?> GetImageAsync(string fileId)
    {
        try
        {
            var fileMeta = await _drive.Files.Get(fileId).ExecuteAsync();
            using var ms = new MemoryStream();
            await _drive.Files.Get(fileId).DownloadAsync(ms);
            return (ms.ToArray(), fileMeta.MimeType ?? "image/jpeg");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to download image {FileId}", fileId);
            return null;
        }
    }

    private async Task<(string? html, string? error)> ReadDocContentAsync(Google.Apis.Drive.v3.Data.File docFile)
    {
        try
        {
            using var ms = new MemoryStream();

            if (docFile.MimeType == "application/vnd.google-apps.document")
            {
                await _drive.Files.Export(docFile.Id,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    .DownloadAsync(ms);
            }
            else
            {
                await _drive.Files.Get(docFile.Id).DownloadAsync(ms);
            }

            ms.Position = 0;
            using var wordDoc = WordprocessingDocument.Open(ms, false);
            var body = wordDoc.MainDocumentPart?.Document?.Body;
            if (body == null)
                return (null, "Документ пустой");

            // Convert to simple HTML
            var html = new System.Text.StringBuilder();
            foreach (var para in body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
            {
                var text = para.InnerText.Trim();
                if (string.IsNullOrWhiteSpace(text))
                {
                    html.AppendLine("<br/>");
                    continue;
                }

                // Check if bold (heading-like)
                var isBold = para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Bold>().Any();
                if (isBold)
                    html.AppendLine($"<p><strong>{System.Net.WebUtility.HtmlEncode(text)}</strong></p>");
                else
                    html.AppendLine($"<p>{System.Net.WebUtility.HtmlEncode(text)}</p>");
            }

            return (html.ToString(), null);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to read doc {Name}", docFile.Name);
            return (null, $"Ошибка чтения файла: {ex.Message}");
        }
    }
}

public class ImageInfo
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string MimeType { get; set; } = "";
    public string? ThumbnailUrl { get; set; }
}
