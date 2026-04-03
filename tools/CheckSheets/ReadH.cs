using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;

var credential = GoogleCredential.FromFile("Pactum.Showcase/google-credentials.json")
    .CreateScoped(SheetsService.Scope.SpreadsheetsReadonly);
var service = new SheetsService(new BaseClientService.Initializer { HttpClientInitializer = credential });

var sid = "15prE_Se2EnRuukWUq7tutPm3S4yKvtirMJvQzp_VFhs";

// Check sheet names
var meta = await service.Spreadsheets.Get(sid).ExecuteAsync();
foreach (var s in meta.Sheets)
    Console.WriteLine($"GID={s.Properties.SheetId} Title=[{s.Properties.Title}] Rows={s.Properties.GridProperties.RowCount}");

// Find sheet by GID 1458578282
var gid = 1458578282;
var sheet = meta.Sheets.FirstOrDefault(s => s.Properties.SheetId == gid);
var sheetName = sheet?.Properties.Title ?? "NOT FOUND";
Console.WriteLine($"\nTarget sheet: [{sheetName}]");

// Read data
var range = $"'{sheetName}'!A1:ZZ";
var resp = await service.Spreadsheets.Values.Get(sid, range).ExecuteAsync();
Console.WriteLine($"Total rows returned: {resp.Values?.Count ?? 0}");
if (resp.Values != null && resp.Values.Count > 0)
{
    Console.WriteLine($"Row 1: {string.Join(" | ", resp.Values[0].Take(5))}");
    if (resp.Values.Count > 1)
        Console.WriteLine($"Row 2: {string.Join(" | ", resp.Values[1].Take(7))}");
    if (resp.Values.Count > 2)
        Console.WriteLine($"Row 3: {string.Join(" | ", resp.Values[2].Take(7))}");
}
