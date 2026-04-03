using Pactum.Showcase.Models;

namespace Pactum.Showcase.Services;

public class GoogleSheetsService
{
    private readonly GoogleSheetsApiService _api;
    private readonly int _headerRow;
    private readonly ILogger<GoogleSheetsService> _logger;

    public GoogleSheetsService(IConfiguration config, ILogger<GoogleSheetsService> logger, GoogleSheetsApiService api)
    {
        _logger = logger;
        _api = api;
        _headerRow = config.GetValue("GoogleSheets:HeaderRow", 2);
    }

    public async Task<List<Business>> GetBusinessesAsync()
    {
        _logger.LogInformation("Fetching sheet via Google API, headerRow={HeaderRow}", _headerRow);

        var allRows = await _api.ReadAllDataAsync();

        _logger.LogInformation("API returned {Count} rows", allRows.Count);

        if (allRows.Count < _headerRow + 1)
        {
            _logger.LogWarning("Sheet has fewer than {Expected} rows, got {Actual}", _headerRow + 1, allRows.Count);
            return [];
        }

        var headers = allRows[_headerRow - 1]
            .Select(h => h?.ToString()?.Trim() ?? "")
            .ToList();

        _logger.LogInformation("Headers ({Count}): {Headers}", headers.Count, string.Join(", ", headers.Take(10)));

        var nameColumnIndex = FindNameColumn(headers);

        var businesses = new List<Business>();
        for (int i = _headerRow; i < allRows.Count; i++)
        {
            var row = allRows[i];
            var properties = new Dictionary<string, string>();

            for (int j = 0; j < headers.Count && j < row.Count; j++)
            {
                var header = headers[j];
                var value = row[j]?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrWhiteSpace(header) && !string.IsNullOrWhiteSpace(value))
                    properties[header] = value;
            }

            if (properties.Count == 0)
                continue;

            var name = nameColumnIndex >= 0 && nameColumnIndex < row.Count
                ? row[nameColumnIndex]?.ToString()?.Trim() ?? ""
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
