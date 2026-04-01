using System.Text.RegularExpressions;
using Pactum.Showcase.Models;

namespace Pactum.Showcase.Services;

public partial class GoogleSheetsService
{
    private readonly string _spreadsheetId;
    private readonly string _gid;
    private readonly int _headerRow;
    private readonly ILogger<GoogleSheetsService> _logger;
    private readonly HttpClient _httpClient;

    public GoogleSheetsService(IConfiguration config, ILogger<GoogleSheetsService> logger, HttpClient httpClient)
    {
        _logger = logger;
        _httpClient = httpClient;
        _spreadsheetId = config["GoogleSheets:SpreadsheetId"]
            ?? throw new InvalidOperationException("GoogleSheets:SpreadsheetId not configured");
        _gid = config["GoogleSheets:Gid"] ?? "0";
        _headerRow = config.GetValue("GoogleSheets:HeaderRow", 2);
    }

    public async Task<List<Business>> GetBusinessesAsync()
    {
        var url = $"https://docs.google.com/spreadsheets/d/{_spreadsheetId}/gviz/tq?tqx=out:csv&gid={_gid}";
        _logger.LogInformation("Fetching sheet from {Url}", url);

        var csv = await _httpClient.GetStringAsync(url);
        var lines = ParseCsvLines(csv);

        if (lines.Count < _headerRow + 1)
        {
            _logger.LogWarning("Sheet has fewer than {Expected} rows, got {Actual}", _headerRow + 1, lines.Count);
            return [];
        }

        var headers = lines[_headerRow - 1].Select(h => h.Trim()).ToList();
        var nameColumnIndex = FindNameColumn(headers);

        var businesses = new List<Business>();
        for (int i = _headerRow; i < lines.Count; i++)
        {
            var row = lines[i];
            var properties = new Dictionary<string, string>();

            for (int j = 0; j < headers.Count && j < row.Count; j++)
            {
                var header = headers[j];
                var value = row[j].Trim();
                if (!string.IsNullOrWhiteSpace(header) && !string.IsNullOrWhiteSpace(value))
                    properties[header] = value;
            }

            if (properties.Count == 0)
                continue;

            var name = nameColumnIndex >= 0 && nameColumnIndex < row.Count
                ? row[nameColumnIndex].Trim()
                : "";

            if (string.IsNullOrWhiteSpace(name))
                name = $"Бизнес #{i}";

            businesses.Add(new Business
            {
                Row = i + 1,
                Name = name,
                Properties = properties
            });
        }

        _logger.LogInformation("Loaded {Count} businesses from Google Sheets", businesses.Count);
        return businesses;
    }

    private static List<List<string>> ParseCsvLines(string csv)
    {
        var results = new List<List<string>>();
        var currentField = new System.Text.StringBuilder();
        var currentRow = new List<string>();
        bool inQuotes = false;

        for (int i = 0; i < csv.Length; i++)
        {
            char c = csv[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < csv.Length && csv[i + 1] == '"')
                    {
                        currentField.Append('"');
                        i++; // skip escaped quote
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    currentField.Append(c);
                }
            }
            else
            {
                if (c == '"')
                {
                    inQuotes = true;
                }
                else if (c == ',')
                {
                    currentRow.Add(currentField.ToString());
                    currentField.Clear();
                }
                else if (c == '\n')
                {
                    currentRow.Add(currentField.ToString());
                    currentField.Clear();
                    results.Add(currentRow);
                    currentRow = new List<string>();
                }
                else if (c == '\r')
                {
                    // skip, \n will handle the line break
                }
                else
                {
                    currentField.Append(c);
                }
            }
        }

        // last row
        if (currentField.Length > 0 || currentRow.Count > 0)
        {
            currentRow.Add(currentField.ToString());
            results.Add(currentRow);
        }

        return results;
    }

    private static int FindNameColumn(List<string> headers)
    {
        var nameKeywords = new[] { "название", "name", "наименование", "компания", "бизнес", "company", "business" };
        for (int i = 0; i < headers.Count; i++)
        {
            var header = headers[i].ToLowerInvariant();
            if (nameKeywords.Any(k => header.Contains(k)))
                return i;
        }
        return 0;
    }
}
