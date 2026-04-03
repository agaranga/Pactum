using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Drive.v3;
using Google.Apis.Services;

var credential = GoogleCredential.FromFile("Pactum.Showcase/google-credentials.json")
    .CreateScoped(SheetsService.Scope.Spreadsheets, DriveService.Scope.DriveReadonly);

var drive = new DriveService(new BaseClientService.Initializer { HttpClientInitializer = credential });
var sheets = new SheetsService(new BaseClientService.Initializer { HttpClientInitializer = credential });

var rootFolderId = "1YOTXieCwwv6NXDJtmZC4jIYFXWql-wIu";
var spreadsheetId = "15prE_Se2EnRuukWUq7tutPm3S4yKvtirMJvQzp_VFhs";
var sheetName = "АнкетаMAX";

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

// List all subfolders
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

// Parse folder names
var folderData = new List<(string id8, string name)>();
foreach (var f in allFolders)
{
    var fname = f.Name.Trim();
    if (fname.Length >= 8 && fname.StartsWith("ID", StringComparison.OrdinalIgnoreCase))
    {
        var id8 = fname[..8];
        var namePart = fname.Length > 9 ? fname[9..].Replace("_", " ").Trim() : "";
        folderData.Add((id8, namePart));
    }
}
Console.WriteLine($"Folders: {folderData.Count}");

// Read spreadsheet
var range = $"'{sheetName}'!B3:G1003";
var response = await sheets.Spreadsheets.Values.Get(spreadsheetId, range).ExecuteAsync();
var rows = response.Values ?? [];
Console.WriteLine($"Sheet rows: {rows.Count}\n");

// STRICT matching only: normalized folder name == normalized sheet name,
// or one is fully contained in the other AND the shorter one is a brand/unique name (>=5 chars, not generic)
var updates = new List<(int row, string id8, string folderName, string sheetTitle)>();
var noMatch = new List<(string id8, string name)>();

foreach (var (id8, folderName) in folderData)
{
    var nFolder = Normalize(folderName);
    int bestRow = -1;
    string bestTitle = "";

    for (int i = 0; i < rows.Count; i++)
    {
        var existingId = rows[i].Count > 0 ? rows[i][0]?.ToString()?.Trim() ?? "" : "";
        var sheetTitle = rows[i].Count > 5 ? rows[i][5]?.ToString()?.Trim() ?? "" : "";
        if (string.IsNullOrWhiteSpace(sheetTitle)) continue;

        var nSheet = Normalize(sheetTitle);

        // Exact match
        if (nFolder == nSheet) { bestRow = i; bestTitle = sheetTitle; break; }

        // Sheet title is a brand name fully contained in folder name
        // e.g. sheet="Velluto", folder="Кафе Velluto на Профсоюзной улице"
        // Skip too-short sheet names (like "Цветы") to avoid false matches
        if (nSheet.Length >= 5 && nFolder.Contains(nSheet) && !IsGeneric(nSheet))
        { bestRow = i; bestTitle = sheetTitle; break; }

        // Folder name fully contained in sheet title
        if (nFolder.Length >= 4 && nSheet.Contains(nFolder) && !IsGeneric(nFolder))
        { bestRow = i; bestTitle = sheetTitle; break; }
    }

    if (bestRow >= 0)
    {
        var rowNum = bestRow + 3;
        var existingId = rows[bestRow].Count > 0 ? rows[bestRow][0]?.ToString()?.Trim() ?? "" : "";
        if (string.IsNullOrWhiteSpace(existingId))
        {
            updates.Add((rowNum, id8, folderName, bestTitle));
            Console.WriteLine($"  MATCH: [{folderName}] -> row {rowNum} [{bestTitle}] = {id8}");
        }
        else
        {
            Console.WriteLine($"  ALREADY: row {rowNum} [{bestTitle}] has [{existingId}]");
        }
    }
    else
    {
        noMatch.Add((id8, folderName));
    }
}

Console.WriteLine($"\nMatched: {updates.Count}");
Console.WriteLine($"No match: {noMatch.Count}");
if (noMatch.Count > 0)
{
    Console.WriteLine("\nUnmatched:");
    foreach (var (id8, name) in noMatch)
        Console.WriteLine($"  {id8} [{name}]");
}

if (args.Contains("--apply") && updates.Count > 0)
{
    var data = updates.Select(u => new ValueRange
    {
        Range = $"'{sheetName}'!B{u.row}",
        Values = [[u.id8]]
    }).ToList();

    var batch = new BatchUpdateValuesRequest { ValueInputOption = "RAW", Data = data };
    var result = await sheets.Spreadsheets.Values.BatchUpdate(batch, spreadsheetId).ExecuteAsync();
    Console.WriteLine($"\nUpdated {result.TotalUpdatedCells} cells!");
}
else if (updates.Count > 0)
{
    Console.WriteLine("\nDry run. Add --apply to write.");
}

static string Normalize(string s) =>
    s.ToLowerInvariant().Replace("ё", "е").Replace("\"", "").Replace("«", "").Replace("»", "")
     .Replace("\n", " ").Replace("\r", "").Trim();

static bool IsGeneric(string normalized)
{
    var generics = new HashSet<string> {
        "кафе", "бар", "пивной бар", "салон", "студия", "клиника", "магазин",
        "производство", "центр", "ремонт", "пункт выдачи", "кофейня",
        "пиццерия", "сауна", "детский центр", "швейное производство",
        "мебельный магазин", "цветочный магазин", "салон красоты",
        "стоматологическая клиника", "центр косметологии",
        "студия маникюра", "автосервис", "пивной паб"
    };
    return generics.Contains(normalized);
}
