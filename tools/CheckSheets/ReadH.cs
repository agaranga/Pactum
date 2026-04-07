using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Docs.v1;
using Google.Apis.Services;

var credential = GoogleCredential.FromFile("Pactum.Showcase/google-credentials.json")
    .CreateScoped(DriveService.Scope.DriveReadonly, DocsService.Scope.DocumentsReadonly);
var drive = new DriveService(new BaseClientService.Initializer { HttpClientInitializer = credential });
var docs = new DocsService(new BaseClientService.Initializer { HttpClientInitializer = credential });

var moscowId = "1gqZY9aJ2CXZ4oJGvl-qb34t3yUngm9Ch";

// Read schemas for the same 5 IDs as sample cards
var ids = new[] { "ID021559", "ID021560", "ID021568" };

foreach (var id in ids)
{
    var req = drive.Files.List();
    req.Q = $"'{moscowId}' in parents and mimeType = 'application/vnd.google-apps.folder' and name contains '{id}'";
    req.Fields = "files(id, name)";
    var folders = await req.ExecuteAsync();
    if (folders.Files.Count == 0) { Console.WriteLine($"No folder for {id}"); continue; }
    var folder = folders.Files[0];

    // Find main doc (schema)
    var docReq = drive.Files.List();
    docReq.Q = $"'{folder.Id}' in parents and (mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' or mimeType = 'application/vnd.google-apps.document')";
    docReq.Fields = "files(id, name, mimeType)";
    var docFiles = await docReq.ExecuteAsync();

    var mainDoc = docFiles.Files.FirstOrDefault(f =>
        f.Name.StartsWith(id, StringComparison.OrdinalIgnoreCase)
        && !f.Name.Contains("card", StringComparison.OrdinalIgnoreCase)
        && !f.Name.Contains("карточка", StringComparison.OrdinalIgnoreCase));

    if (mainDoc == null) { Console.WriteLine($"No schema for {id}"); continue; }

    Console.WriteLine($"\n=== SCHEMA: {mainDoc.Name} [{mainDoc.MimeType}] ===\n");

    if (mainDoc.MimeType == "application/vnd.google-apps.document")
    {
        var doc = await docs.Documents.Get(mainDoc.Id).ExecuteAsync();
        foreach (var el in doc.Body.Content)
        {
            if (el.Paragraph != null)
            {
                var text = string.Join("", el.Paragraph.Elements?
                    .Where(e => e.TextRun != null)
                    .Select(e => e.TextRun.Content) ?? []).TrimEnd('\n');
                if (!string.IsNullOrWhiteSpace(text))
                    Console.WriteLine(text);
            }
            else if (el.Table != null)
            {
                foreach (var row in el.Table.TableRows)
                {
                    var cells = row.TableCells.Select(c =>
                        string.Join(" ", c.Content
                            .Where(ce => ce.Paragraph != null)
                            .Select(ce => string.Join("", ce.Paragraph.Elements?
                                .Where(e => e.TextRun != null)
                                .Select(e => e.TextRun.Content?.Replace("\n", " ").Trim()) ?? []))).Trim());
                    Console.WriteLine($"| {string.Join(" | ", cells)} |");
                }
            }
        }
    }
    else
    {
        using var ms = new MemoryStream();
        await drive.Files.Get(mainDoc.Id).DownloadAsync(ms);
        ms.Position = 0;
        using var wordDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
        var body = wordDoc.MainDocumentPart?.Document?.Body;
        if (body == null) continue;
        foreach (var child in body.ChildElements)
        {
            if (child is DocumentFormat.OpenXml.Wordprocessing.Paragraph para)
            {
                var text = para.InnerText.Trim();
                if (!string.IsNullOrWhiteSpace(text)) Console.WriteLine(text);
            }
            else if (child is DocumentFormat.OpenXml.Wordprocessing.Table table)
            {
                foreach (var row in table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>())
                {
                    var cells = row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>()
                        .Select(c => c.InnerText.Trim());
                    Console.WriteLine($"| {string.Join(" | ", cells)} |");
                }
            }
        }
    }

    Console.WriteLine("\n========================================");
}
