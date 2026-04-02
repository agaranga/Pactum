using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;

namespace Pactum.Showcase.Services;

/// <summary>
/// Authenticated Google Sheets service for read/write operations via API.
/// </summary>
public class GoogleSheetsApiService
{
    private readonly SheetsService _sheets;
    private readonly string _spreadsheetId;
    private readonly string _gid;
    private readonly int _headerRow;
    private readonly ILogger<GoogleSheetsApiService> _logger;

    public GoogleSheetsApiService(IConfiguration config, ILogger<GoogleSheetsApiService> logger)
    {
        _logger = logger;
        _spreadsheetId = config["GoogleSheets:SpreadsheetId"]
            ?? throw new InvalidOperationException("GoogleSheets:SpreadsheetId not configured");
        _gid = config["GoogleSheets:Gid"] ?? "0";
        _headerRow = config.GetValue("GoogleSheets:HeaderRow", 2);

        var credPath = config["GoogleSheets:CredentialsPath"];
        var credJson = config["GoogleSheets:CredentialsJson"];

        GoogleCredential credential;
        if (!string.IsNullOrWhiteSpace(credJson))
        {
            credential = GoogleCredential.FromJson(credJson);
        }
        else if (!string.IsNullOrWhiteSpace(credPath) && File.Exists(credPath))
        {
            credential = GoogleCredential.FromFile(credPath);
        }
        else
        {
            throw new InvalidOperationException(
                "Configure GoogleSheets:CredentialsPath or GoogleSheets:CredentialsJson for Google API access");
        }

        credential = credential.CreateScoped(SheetsService.Scope.Spreadsheets);
        _sheets = new SheetsService(new BaseClientService.Initializer { HttpClientInitializer = credential });
    }

    /// <summary>
    /// Get the sheet name by GID.
    /// </summary>
    public async Task<string> GetSheetNameAsync()
    {
        var meta = await _sheets.Spreadsheets.Get(_spreadsheetId).ExecuteAsync();
        var gidInt = int.Parse(_gid);
        var sheet = meta.Sheets.FirstOrDefault(s => s.Properties.SheetId == gidInt);
        return sheet?.Properties.Title ?? "Sheet1";
    }

    /// <summary>
    /// Clear all data rows (keep headers).
    /// </summary>
    public async Task<int> ClearDataAsync()
    {
        var sheetName = await GetSheetNameAsync();
        var dataStart = _headerRow + 1;
        var range = $"'{sheetName}'!A{dataStart}:ZZ";

        _logger.LogInformation("Clearing data range {Range}", range);

        var request = _sheets.Spreadsheets.Values.Clear(new ClearValuesRequest(), _spreadsheetId, range);
        var result = await request.ExecuteAsync();

        _logger.LogInformation("Cleared range {Range}", result.ClearedRange);
        return dataStart;
    }

    /// <summary>
    /// Write rows of data starting from the given row.
    /// </summary>
    public async Task<int> WriteDataAsync(IList<IList<object>> rows, int startRow)
    {
        var sheetName = await GetSheetNameAsync();
        var range = $"'{sheetName}'!A{startRow}";

        _logger.LogInformation("Writing {Count} rows starting at {Range}", rows.Count, range);

        var body = new ValueRange { Values = rows };
        var request = _sheets.Spreadsheets.Values.Update(body, _spreadsheetId, range);
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

        var result = await request.ExecuteAsync();
        _logger.LogInformation("Updated {Cells} cells", result.UpdatedCells);
        return result.UpdatedCells ?? 0;
    }

    /// <summary>
    /// Read all data rows (after header).
    /// </summary>
    public async Task<IList<IList<object>>> ReadAllDataAsync()
    {
        var sheetName = await GetSheetNameAsync();
        var range = $"'{sheetName}'!A1:ZZ";

        var response = await _sheets.Spreadsheets.Values.Get(_spreadsheetId, range).ExecuteAsync();
        return response.Values ?? new List<IList<object>>();
    }

    /// <summary>
    /// Extract cities from address column (I) and write to city column (H).
    /// </summary>
    public async Task<(int updated, int skipped)> ExtractCitiesAsync()
    {
        var sheetName = await GetSheetNameAsync();
        var dataStart = _headerRow + 1;
        var range = $"'{sheetName}'!H{dataStart}:I1003";

        var response = await _sheets.Spreadsheets.Values.Get(_spreadsheetId, range).ExecuteAsync();
        var rows = response.Values ?? [];

        var updates = new List<ValueRange>();
        int updated = 0, skipped = 0;

        for (int i = 0; i < rows.Count; i++)
        {
            var rowNum = i + dataStart;
            var existingCity = rows[i].Count > 0 ? rows[i][0]?.ToString()?.Trim() ?? "" : "";
            var address = rows[i].Count > 1 ? rows[i][1]?.ToString()?.Trim() ?? "" : "";

            if (!string.IsNullOrWhiteSpace(existingCity) || string.IsNullOrWhiteSpace(address))
                continue;

            var city = ExtractCity(address);
            if (city != null)
            {
                updates.Add(new ValueRange
                {
                    Range = $"'{sheetName}'!H{rowNum}",
                    Values = [[city]]
                });
                updated++;
            }
            else
            {
                skipped++;
            }
        }

        if (updates.Count > 0)
        {
            var batch = new BatchUpdateValuesRequest
            {
                ValueInputOption = "RAW",
                Data = updates
            };
            await _sheets.Spreadsheets.Values.BatchUpdate(batch, _spreadsheetId).ExecuteAsync();
        }

        _logger.LogInformation("Cities: updated {Updated}, skipped {Skipped}", updated, skipped);
        return (updated, skipped);
    }

    private static readonly HashSet<string> KnownCities = new(StringComparer.OrdinalIgnoreCase)
    {
        "Москва", "Санкт-Петербург", "Санкт Петербург", "СПб",
        "Новосибирск", "Екатеринбург", "Казань", "Нижний Новгород",
        "Челябинск", "Самара", "Омск", "Ростов-на-Дону",
        "Уфа", "Красноярск", "Воронеж", "Пермь", "Волгоград",
        "Краснодар", "Саратов", "Тюмень", "Тольятти", "Ижевск",
        "Барнаул", "Ульяновск", "Иркутск", "Хабаровск", "Ярославль",
        "Владивосток", "Махачкала", "Томск", "Оренбург", "Кемерово",
        "Рязань", "Астрахань", "Пенза", "Липецк", "Тула", "Киров",
        "Калининград", "Курск", "Сочи", "Ставрополь", "Тверь",
        "Брянск", "Иваново", "Белгород", "Сургут",
        "Владимир", "Архангельск", "Калуга", "Смоленск",
        "Подольск", "Мытищи", "Балашиха", "Химки", "Люберцы",
        "Королёв", "Одинцово", "Домодедово", "Зеленоград",
        "Долгопрудный", "Красногорск", "Щёлково", "Реутов",
        "Котельники", "Видное", "Лобня", "Дзержинский",
        "Серпухов", "Обнинск", "Пушкино", "Жуковский",
        "Раменское", "Ногинск", "Коломна", "Электросталь",
        "Чехов", "Клин", "Дубна", "Наро-Фоминск",
    };

    private static string? ExtractCity(string address)
    {
        if (string.IsNullOrWhiteSpace(address)) return null;

        var parts = address.Split(',').Select(p => p.Trim()).Where(p => p.Length > 0).ToList();
        if (parts.Count == 0) return null;

        var first = parts[0];

        if (first.Contains("обл", StringComparison.OrdinalIgnoreCase))
        {
            if (parts.Count >= 2)
            {
                var candidate = parts[1].Trim();
                if (KnownCities.Contains(candidate)) return candidate;
                if (candidate.Length >= 3 && candidate.Length <= 30
                    && !candidate.Contains("ул") && !candidate.Contains("пр"))
                    return candidate;
            }
            return "Московская обл.";
        }

        if (KnownCities.Contains(first)) return first;

        return null;
    }
}
