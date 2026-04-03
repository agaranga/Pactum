using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

var credential = GoogleCredential.FromFile("Pactum.Showcase/google-credentials.json")
    .CreateScoped(SheetsService.Scope.SpreadsheetsReadonly, DriveService.Scope.DriveReadonly);

var drive = new DriveService(new BaseClientService.Initializer { HttpClientInitializer = credential });
var sheets = new SheetsService(new BaseClientService.Initializer { HttpClientInitializer = credential });

var rootFolderId = "1YOTXieCwwv6NXDJtmZC4jIYFXWql-wIu";
var spreadsheetId = "15prE_Se2EnRuukWUq7tutPm3S4yKvtirMJvQzp_VFhs";

// Find Москва folder
var rootReq = drive.Files.List();
rootReq.Q = $"'{rootFolderId}' in parents and mimeType = 'application/vnd.google-apps.folder'";
rootReq.Fields = "files(id, name)";
var rootFolders = await rootReq.ExecuteAsync();
var bazaFolder = rootFolders.Files.FirstOrDefault(f => f.Name.Contains("База"));
Google.Apis.Drive.v3.Data.File? moscowFolder;
if (bazaFolder != null)
{
    var sub = drive.Files.List();
    sub.Q = $"'{bazaFolder.Id}' in parents and mimeType = 'application/vnd.google-apps.folder'";
    sub.Fields = "files(id, name)";
    moscowFolder = (await sub.ExecuteAsync()).Files.FirstOrDefault(f => f.Name.Contains("Москва"));
}
else
{
    moscowFolder = rootFolders.Files.FirstOrDefault(f => f.Name.Contains("Москва"));
}
if (moscowFolder == null) { Console.WriteLine("ERROR: Москва not found!"); return; }

// Get all subfolders
var allFolders = new List<Google.Apis.Drive.v3.Data.File>();
string? pageToken = null;
do
{
    var listReq = drive.Files.List();
    listReq.Q = $"'{moscowFolder.Id}' in parents and mimeType = 'application/vnd.google-apps.folder'";
    listReq.Fields = "nextPageToken, files(id, name)";
    listReq.PageSize = 1000;
    listReq.PageToken = pageToken;
    var page = await listReq.ExecuteAsync();
    allFolders.AddRange(page.Files);
    pageToken = page.NextPageToken;
} while (pageToken != null);

// Build folder map: id8 -> folderId
var folderMap = new Dictionary<string, (string folderId, string folderName)>(StringComparer.OrdinalIgnoreCase);
foreach (var f in allFolders)
{
    var fname = f.Name.Trim();
    if (fname.Length >= 8 && fname.StartsWith("ID", StringComparison.OrdinalIgnoreCase))
        folderMap[fname[..8]] = (f.Id, fname);
}

// Read spreadsheet rows with IDs
var resp = await sheets.Spreadsheets.Values.Get(spreadsheetId, "'АнкетаMAX'!A3:I1003").ExecuteAsync();
var rows = resp.Values ?? [];

int verified = 0, mismatch = 0, noFolder = 0, noDoc = 0;

for (int i = 0; i < rows.Count; i++)
{
    var rowNum = i + 3;
    var idCell = rows[i].Count > 1 ? rows[i][1]?.ToString()?.Trim() ?? "" : "";
    if (string.IsNullOrWhiteSpace(idCell) || !idCell.StartsWith("ID", StringComparison.OrdinalIgnoreCase))
        continue;

    var sheetName = rows[i].Count > 6 ? rows[i][6]?.ToString()?.Trim() ?? "" : "";  // G
    var sheetAddr = rows[i].Count > 8 ? rows[i][8]?.ToString()?.Trim() ?? "" : "";  // I

    if (!folderMap.TryGetValue(idCell, out var folder))
    {
        Console.WriteLine($"Row {rowNum} [{idCell}]: FOLDER NOT FOUND");
        noFolder++;
        continue;
    }

    // Find docx or doc files in folder
    var docReq = drive.Files.List();
    docReq.Q = $"'{folder.folderId}' in parents and (mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' or mimeType = 'application/msword' or mimeType = 'application/vnd.google-apps.document')";
    docReq.Fields = "files(id, name, mimeType)";
    var docFiles = await docReq.ExecuteAsync();

    if (docFiles.Files.Count == 0)
    {
        Console.WriteLine($"Row {rowNum} [{idCell}] [{sheetName}]: NO DOC in folder [{folder.folderName}]");
        noDoc++;
        continue;
    }

    // Download and read docx
    var docFile = docFiles.Files[0];
    string docText = "";

    try
    {
        using var ms = new MemoryStream();
        if (docFile.MimeType == "application/vnd.google-apps.document")
        {
            // Export Google Doc as docx
            await drive.Files.Export(docFile.Id, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                .DownloadAsync(ms);
        }
        else
        {
            // Download Word file directly
            await drive.Files.Get(docFile.Id).DownloadAsync(ms);
        }

        ms.Position = 0;
        using var wordDoc = WordprocessingDocument.Open(ms, false);
        var body = wordDoc.MainDocumentPart?.Document?.Body;
        if (body != null)
            docText = body.InnerText;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Row {rowNum} [{idCell}] [{sheetName}]: ERROR reading doc: {ex.Message}");
        mismatch++;
        continue;
    }

    var docLower = docText.ToLowerInvariant().Replace("ё", "е");
    var nameNorm = sheetName.ToLowerInvariant().Replace("ё", "е").Replace("\"", "").Replace("«", "").Replace("»", "").Trim();

    // Check name: at least one significant word (>=4 chars) from sheet name found in doc
    var nameWords = nameNorm.Split(new[] { ' ', ',', '-', '.' }, StringSplitOptions.RemoveEmptyEntries)
        .Where(w => w.Length >= 4).ToList();
    var nameMatches = nameWords.Count(w => docLower.Contains(w));
    var nameOk = nameWords.Count == 0 || nameMatches > 0;

    // Check address
    var addrNorm = sheetAddr.ToLowerInvariant().Replace("ё", "е");
    var addrWords = addrNorm.Split(new[] { ' ', ',', '.' }, StringSplitOptions.RemoveEmptyEntries)
        .Where(w => w.Length >= 4).ToList();
    var addrMatches = addrWords.Count(w => docLower.Contains(w));
    var addrOk = addrWords.Count == 0 || addrMatches > 0;

    var status = (nameOk && addrOk) ? "OK" : "MISMATCH";
    Console.WriteLine($"Row {rowNum} [{idCell}] [{sheetName}]: {status} (name: {nameMatches}/{nameWords.Count}, addr: {addrMatches}/{addrWords.Count})");

    if (status == "MISMATCH")
    {
        Console.WriteLine($"  Sheet name: [{sheetName}]");
        Console.WriteLine($"  Sheet addr: [{sheetAddr}]");
        Console.WriteLine($"  Doc file: [{docFile.Name}]");
        Console.WriteLine($"  Doc preview: [{docText[..Math.Min(300, docText.Length)]}]");
        mismatch++;
    }
    else
    {
        verified++;
    }
}

Console.WriteLine($"\n=== Summary ===");
Console.WriteLine($"Verified OK: {verified}");
Console.WriteLine($"Mismatch: {mismatch}");
Console.WriteLine($"No folder: {noFolder}");
Console.WriteLine($"No doc file: {noDoc}");
