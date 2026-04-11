using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Docs.v1;
using Google.Apis.Services;
using DocumentFormat.OpenXml.Packaging;

namespace Pactum.Showcase.Services;

public class DriveFileService
{
    private readonly DriveService _drive;
    private readonly DocsService _docs;
    private readonly GoogleOAuthService _oauth;
    private readonly string _rootFolderId;
    private readonly ILogger<DriveFileService> _logger;

    // city -> folderId from config
    private readonly Dictionary<string, string> _cityFolderIds = new(StringComparer.OrdinalIgnoreCase);
    // Cache: externalId -> folderId
    private readonly Dictionary<string, string> _bizFolderIds = new(StringComparer.OrdinalIgnoreCase);
    private readonly SemaphoreSlim _lock = new(1, 1);

    public DriveFileService(IConfiguration config, ILogger<DriveFileService> logger, GoogleOAuthService oauth)
    {
        _logger = logger;
        _oauth = oauth;
        _rootFolderId = config["GoogleDrive:RootFolderId"] ?? "1WHH6FCHPNyn09OZAW1Y9wL-Rdn2OrofO";

        // Load city folder IDs from config
        var cityFolders = config.GetSection("GoogleDrive:CityFolders");
        foreach (var child in cityFolders.GetChildren())
        {
            if (!string.IsNullOrWhiteSpace(child.Value))
                _cityFolderIds[child.Key] = child.Value;
        }

        var credPath = config["GoogleSheets:CredentialsPath"];
        var credJson = config["GoogleSheets:CredentialsJson"];

        GoogleCredential credential;
        if (!string.IsNullOrWhiteSpace(credJson))
            credential = GoogleCredential.FromJson(credJson);
        else if (!string.IsNullOrWhiteSpace(credPath) && File.Exists(credPath))
            credential = GoogleCredential.FromFile(credPath);
        else
            throw new InvalidOperationException("Google credentials not configured");

        credential = credential.CreateScoped(DriveService.Scope.Drive, DocsService.Scope.Documents);
        var init = new BaseClientService.Initializer { HttpClientInitializer = credential };
        _drive = new DriveService(init);
        _docs = new DocsService(init);
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

            // Search only in the specified city folder
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
        var diag = new System.Text.StringBuilder();
        diag.AppendLine($"<hr/><small class='text-muted'><b>Диагностика:</b><br/>");
        diag.AppendLine($"ExternalId: {externalId}<br/>");

        var folderId = await FindBizFolderAsync(externalId);
        diag.AppendLine($"FolderId: {folderId ?? "NOT FOUND"}<br/>");
        if (folderId == null)
        {
            // Show what city folders we have
            diag.AppendLine($"City folders cached: {string.Join(", ", _cityFolderIds.Keys)}<br/>");
            diag.AppendLine("</small>");
            return (null, "Папка не найдена на Google Drive" + diag);
        }

        // List ALL files in folder for diagnostics
        var allReq = _drive.Files.List();
        allReq.Q = $"'{folderId}' in parents";
        allReq.Fields = "files(id, name, mimeType)";
        var allFiles = await allReq.ExecuteAsync();
        diag.AppendLine($"Файлов в папке: {allFiles.Files.Count}<br/>");
        foreach (var f in allFiles.Files)
            diag.AppendLine($"  - {f.Name} [{f.MimeType}]<br/>");

        // Find doc files
        var docFiles = allFiles.Files.Where(f =>
            f.MimeType == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            || f.MimeType == "application/msword"
            || f.MimeType == "application/vnd.google-apps.document").ToList();

        diag.AppendLine($"Документов: {docFiles.Count}<br/>");

        var mainDoc = docFiles.FirstOrDefault(f =>
            f.Name.StartsWith(externalId, StringComparison.OrdinalIgnoreCase)
            && !f.Name.Contains("карточка", StringComparison.OrdinalIgnoreCase)
            && !f.Name.Contains("_card", StringComparison.OrdinalIgnoreCase));

        diag.AppendLine($"Main doc match: {mainDoc?.Name ?? "NOT FOUND"}<br/>");
        diag.AppendLine("</small>");

        if (mainDoc == null)
            return (null, "Файл описания не найден" + diag);

        var (html, error) = await ReadDocContentAsync(mainDoc);
        if (error != null)
            return (null, error + diag);

        return (html + diag, null);
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

        // Look for ID_card.docx or legacy ID_карточка
        var cardDoc = files.Files.FirstOrDefault(f =>
            f.Name.Contains("_card", StringComparison.OrdinalIgnoreCase))
            ?? files.Files.FirstOrDefault(f =>
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

    /// <summary>
    /// Check what files exist in the business folder.
    /// </summary>
    public async Task<FilePresenceInfo> CheckFilesAsync(string externalId)
    {
        var info = new FilePresenceInfo();
        var folderId = await FindBizFolderAsync(externalId);
        if (folderId == null)
            return info;

        info.HasFolder = true;

        var req = _drive.Files.List();
        req.Q = $"'{folderId}' in parents";
        req.Fields = "files(name, mimeType)";
        var result = await req.ExecuteAsync();

        foreach (var f in result.Files)
        {
            var isDoc = f.MimeType is "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                or "application/msword" or "application/vnd.google-apps.document";

            if (isDoc && f.Name.StartsWith(externalId, StringComparison.OrdinalIgnoreCase)
                && !f.Name.Contains("карточка", StringComparison.OrdinalIgnoreCase))
            {
                info.HasMainDoc = true;
                info.MainDocType = f.MimeType == "application/vnd.google-apps.document" ? "Google Doc" : "Word";
            }

            if (isDoc && (f.Name.Contains("_card", StringComparison.OrdinalIgnoreCase)
                || f.Name.Contains("карточка", StringComparison.OrdinalIgnoreCase)))
                info.HasCardDoc = true;

            if (f.MimeType?.StartsWith("image/") == true)
                info.HasImages = true;
        }

        return info;
    }

    private async Task<(string? html, string? error)> ReadDocContentAsync(Google.Apis.Drive.v3.Data.File docFile)
    {
        try
        {
            if (docFile.MimeType == "application/vnd.google-apps.document")
                return await ReadGoogleDocAsync(docFile.Id);
            else
                return await ReadWordDocAsync(docFile.Id);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to read doc {Name}", docFile.Name);
            return (null, $"Ошибка чтения файла: {ex.Message}");
        }
    }

    private async Task<(string? html, string? error)> ReadGoogleDocAsync(string docId)
    {
        var doc = await _docs.Documents.Get(docId).ExecuteAsync();
        if (doc.Body?.Content == null)
            return (null, "Документ пустой");

        var html = new System.Text.StringBuilder();

        foreach (var element in doc.Body.Content)
        {
            if (element.Paragraph != null)
            {
                var parts = new List<string>();
                bool anyBold = false;
                foreach (var pe in element.Paragraph.Elements ?? [])
                {
                    if (pe.TextRun == null) continue;
                    var text = pe.TextRun.Content?.Replace("\v", "\n").Replace("\x0B", "\n");
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    var isBold = pe.TextRun.TextStyle?.Bold == true;
                    if (isBold) anyBold = true;
                    parts.Add(isBold ? $"<strong>{Enc(text)}</strong>" : Enc(text));
                }
                var content = string.Join("", parts).Trim();
                if (string.IsNullOrWhiteSpace(content)) continue;

                // Replace newlines with <br/>
                content = content.Replace("\n", "<br/>");

                if (anyBold && content.Length < 100)
                    html.AppendLine($"<h6>{content}</h6>");
                else
                    html.AppendLine($"<p>{content}</p>");
            }
            else if (element.Table != null)
            {
                html.AppendLine("<table class='table table-sm table-bordered mb-3'>");
                bool firstRow = true;
                foreach (var row in element.Table.TableRows)
                {
                    html.AppendLine("<tr>");
                    bool isFirstCol = true;
                    foreach (var cell in row.TableCells)
                    {
                        var cellParts = new List<string>();
                        foreach (var ce in cell.Content)
                        {
                            if (ce.Paragraph == null) continue;
                            var paraTexts = new List<string>();
                            foreach (var pe in ce.Paragraph.Elements ?? [])
                            {
                                if (pe.TextRun == null) continue;
                                var text = pe.TextRun.Content?.TrimEnd('\n')
                                    .Replace("\v", "\n").Replace("\x0B", "\n") ?? "";
                                if (string.IsNullOrWhiteSpace(text)) continue;
                                var isBold = pe.TextRun.TextStyle?.Bold == true;
                                paraTexts.Add(isBold
                                    ? $"<strong>{Enc(text).Replace("\n", "<br/>")}</strong>"
                                    : Enc(text).Replace("\n", "<br/>"));
                            }
                            var paraContent = string.Join("", paraTexts).Trim();
                            if (!string.IsNullOrWhiteSpace(paraContent))
                                cellParts.Add(paraContent);
                        }
                        var cellHtml = string.Join("<br/>", cellParts);

                        if (firstRow)
                            html.AppendLine($"<th>{cellHtml}</th>");
                        else if (isFirstCol)
                            html.AppendLine($"<td class='text-muted fw-semibold' style='width:35%'>{cellHtml}</td>");
                        else
                            html.AppendLine($"<td>{cellHtml}</td>");
                        isFirstCol = false;
                    }
                    html.AppendLine("</tr>");
                    firstRow = false;
                }
                html.AppendLine("</table>");
            }
        }

        return (html.ToString(), null);
    }

    private async Task<(string? html, string? error)> ReadWordDocAsync(string fileId)
    {
        using var ms = new MemoryStream();
        await _drive.Files.Get(fileId).DownloadAsync(ms);
        ms.Position = 0;

        using var wordDoc = WordprocessingDocument.Open(ms, false);
        var body = wordDoc.MainDocumentPart?.Document?.Body;
        if (body == null)
            return (null, "Документ пустой");

        var html = new System.Text.StringBuilder();
        foreach (var child in body.ChildElements)
        {
            if (child is DocumentFormat.OpenXml.Wordprocessing.Paragraph para)
            {
                var text = para.InnerText.Trim();
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                var isBold = para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Bold>().Any();
                if (isBold)
                    html.AppendLine($"<h6>{Enc(text)}</h6>");
                else
                    html.AppendLine($"<p>{Enc(text)}</p>");
            }
            else if (child is DocumentFormat.OpenXml.Wordprocessing.Table table)
            {
                html.AppendLine("<table class='table table-sm table-bordered mb-3'>");
                bool firstRow = true;
                foreach (var row in table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>())
                {
                    html.AppendLine("<tr>");
                    bool isFirstCol = true;
                    foreach (var cell in row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>())
                    {
                        var cellHtml = WordCellToHtml(cell);
                        if (firstRow)
                            html.AppendLine($"<th>{cellHtml}</th>");
                        else if (isFirstCol)
                            html.AppendLine($"<td class='text-muted fw-semibold' style='width:35%'>{cellHtml}</td>");
                        else
                            html.AppendLine($"<td>{cellHtml}</td>");
                        isFirstCol = false;
                    }
                    html.AppendLine("</tr>");
                    firstRow = false;
                }
                html.AppendLine("</table>");
            }
        }

        return (html.ToString(), null);
    }

    /// <summary>
    /// Save generated card as a Google Doc in the business folder on Drive.
    /// </summary>
    public async Task<(bool success, string? error)> SaveCardDocxAsync(string externalId, string markdownText)
    {
        var oauthDrive = await _oauth.GetDriveServiceAsync();
        if (oauthDrive == null)
            return (false, "Google OAuth не настроен. Перейдите в Админ → Подключить Google Drive.");

        var folderId = await FindBizFolderAsync(externalId);
        if (folderId == null)
            return (false, "Папка не найдена на Google Drive");

        var fileName = $"{externalId}_card.docx";

        // Delete existing card file if present
        var existingReq = _drive.Files.List();
        existingReq.Q = $"'{folderId}' in parents and (name contains '{externalId}_card')";
        existingReq.Fields = "files(id, name)";
        var existing = await existingReq.ExecuteAsync();
        foreach (var old in existing.Files)
        {
            try { await oauthDrive.Files.Delete(old.Id).ExecuteAsync(); }
            catch { }
        }

        // Build docx in memory
        using var ms = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
            var body = mainPart.Document.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Body());

            foreach (var line in markdownText.Split('\n'))
            {
                var trimmed = line.Trim();
                if (string.IsNullOrWhiteSpace(trimmed))
                {
                    body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                    continue;
                }

                var para = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                var segments = ParseBoldSegments(trimmed);
                foreach (var (text, isBold) in segments)
                {
                    var run = new DocumentFormat.OpenXml.Wordprocessing.Run();
                    if (isBold)
                        run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.RunProperties(
                            new DocumentFormat.OpenXml.Wordprocessing.Bold()));
                    run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(text)
                        { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
                    para.AppendChild(run);
                }
                body.AppendChild(para);
            }
        }

        ms.Position = 0;

        // Upload docx via OAuth
        var fileMetadata = new Google.Apis.Drive.v3.Data.File
        {
            Name = fileName,
            Parents = [folderId],
            MimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        };

        var upload = oauthDrive.Files.Create(fileMetadata, ms,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        upload.Fields = "id";
        var result = await upload.UploadAsync();

        if (result.Status == Google.Apis.Upload.UploadStatus.Completed)
        {
            _logger.LogInformation("Card saved: {FileName} in folder {FolderId}", fileName, folderId);
            return (true, null);
        }

        return (false, $"Upload failed: {result.Exception?.Message}");
    }

    private static List<(string text, bool bold)> ParseBoldSegments(string line)
    {
        var result = new List<(string text, bool bold)>();
        var parts = line.Split("**");
        bool isBold = false;
        foreach (var part in parts)
        {
            if (!string.IsNullOrEmpty(part))
                result.Add((part, isBold));
            isBold = !isBold;
        }
        return result;
    }

    private static string Enc(string s) => System.Net.WebUtility.HtmlEncode(s);

    private static string WordCellToHtml(DocumentFormat.OpenXml.Wordprocessing.TableCell cell)
    {
        var parts = new List<string>();
        foreach (var p in cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
        {
            var text = p.InnerText.Trim();
            if (!string.IsNullOrWhiteSpace(text))
            {
                var isBold = p.Descendants<DocumentFormat.OpenXml.Wordprocessing.Bold>().Any();
                parts.Add(isBold ? $"<strong>{Enc(text)}</strong>" : Enc(text));
            }
        }
        return string.Join("<br/>", parts);
    }
}

public class FilePresenceInfo
{
    public bool HasFolder { get; set; }
    public bool HasMainDoc { get; set; }
    public bool HasCardDoc { get; set; }
    public bool HasImages { get; set; }
    public string MainDocType { get; set; } = ""; // "Word" or "Google Doc"
}

public class ImageInfo
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string MimeType { get; set; } = "";
    public string? ThumbnailUrl { get; set; }
}
