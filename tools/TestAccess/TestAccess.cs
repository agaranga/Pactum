using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Drive.v3;
using Google.Apis.Services;

var credPath = args.Length > 0 ? args[0] : "Pactum.Showcase/google-credentials.json";
var credential = GoogleCredential.FromFile(credPath)
    .CreateScoped(SheetsService.Scope.Spreadsheets, DriveService.Scope.Drive);

// Test Sheets
var sheetsService = new SheetsService(new BaseClientService.Initializer { HttpClientInitializer = credential });
var spreadsheetId = "1-S7T8jYiGduIw3YqYGAji2GgPipkLOOAfzrmZ3CZkMM";
try {
    var meta = await sheetsService.Spreadsheets.Get(spreadsheetId).ExecuteAsync();
    Console.WriteLine($"[Sheets OK] Sheets count: {meta.Sheets.Count}");
} catch (Exception ex) {
    Console.WriteLine($"[Sheets FAIL] {ex.Message}");
}

// Test Drive
var driveService = new DriveService(new BaseClientService.Initializer { HttpClientInitializer = credential });
var folderId = "1YOTXieCwwv6NXDJtmZC4jIYFXWql-wIu";
try {
    var req = driveService.Files.List();
    req.Q = $"'{folderId}' in parents";
    req.Fields = "files(id, name, mimeType)";
    var result = await req.ExecuteAsync();
    Console.WriteLine($"[Drive OK] Files in folder: {result.Files.Count}");
    foreach (var f in result.Files)
        Console.WriteLine($"  - {f.Name} ({f.MimeType})");
} catch (Exception ex) {
    Console.WriteLine($"[Drive FAIL] {ex.Message}");
}
