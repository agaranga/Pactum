using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;

var credPath = "Pactum.Showcase/google-credentials.json";
var spreadsheetId = "1KvX2qEqp6146jIeUhJflJMxZbkjC8FgRnb6j8ZkExIA";
var sheetName = "АнкетаMAX";

// Known cities for validation
var knownCities = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
{
    "Москва", "Санкт-Петербург", "Санкт Петербург", "СПб",
    "Новосибирск", "Екатеринбург", "Казань", "Нижний Новгород",
    "Челябинск", "Самара", "Омск", "Ростов-на-Дону",
    "Уфа", "Красноярск", "Воронеж", "Пермь", "Волгоград",
    "Краснодар", "Саратов", "Тюмень", "Тольятти", "Ижевск",
    "Барнаул", "Ульяновск", "Иркутск", "Хабаровск", "Ярославль",
    "Владивосток", "Махачкала", "Томск", "Оренбург", "Кемерово",
    "Новокузнецк", "Рязань", "Астрахань", "Набережные Челны",
    "Пенза", "Липецк", "Тула", "Киров", "Чебоксары",
    "Калининград", "Курск", "Сочи", "Ставрополь", "Тверь",
    "Брянск", "Иваново", "Белгород", "Сургут",
    "Владимир", "Архангельск", "Калуга", "Смоленск",
    "Подольск", "Мытищи", "Балашиха", "Химки", "Люберцы",
    "Королёв", "Одинцово", "Домодедово", "Зеленоград",
    "Долгопрудный", "Красногорск", "Щёлково", "Реутов",
    "Котельники", "Видное", "Лобня", "Дзержинский",
    "Серпухов", "Обнинск", "Пушкино", "Жуковский",
    "Раменское", "Ногинск", "Коломна", "Электросталь",
    "Орехово-Зуево", "Дубна", "Клин", "Воскресенск",
    "Чехов", "Егорьевск", "Наро-Фоминск", "Павловский Посад",
};

var credential = GoogleCredential.FromFile(credPath)
    .CreateScoped(SheetsService.Scope.Spreadsheets);

var service = new SheetsService(new BaseClientService.Initializer { HttpClientInitializer = credential });

// Read columns H and I (rows 3+)
var range = $"'{sheetName}'!H3:I1003";
var response = await service.Spreadsheets.Values.Get(spreadsheetId, range).ExecuteAsync();
var rows = response.Values ?? [];

Console.WriteLine($"Read {rows.Count} rows");

var updates = new List<(int row, string city)>();
var skipped = new List<(int row, string addr)>();

for (int i = 0; i < rows.Count; i++)
{
    var rowNum = i + 3; // 1-based, starting from row 3
    var existingCity = rows[i].Count > 0 ? rows[i][0]?.ToString()?.Trim() ?? "" : "";
    var address = rows[i].Count > 1 ? rows[i][1]?.ToString()?.Trim() ?? "" : "";

    // Skip if city already filled or no address
    if (!string.IsNullOrWhiteSpace(existingCity) || string.IsNullOrWhiteSpace(address))
        continue;

    var city = ExtractCity(address);
    if (city != null)
    {
        updates.Add((rowNum, city));
        Console.WriteLine($"  Row {rowNum}: [{address}] -> [{city}]");
    }
    else
    {
        skipped.Add((rowNum, address));
    }
}

Console.WriteLine($"\nWill update {updates.Count} rows, skipped {skipped.Count}");

if (skipped.Count > 0)
{
    Console.WriteLine("\nSkipped (could not extract city):");
    foreach (var (row, addr) in skipped.Take(20))
        Console.WriteLine($"  Row {row}: [{addr}]");
}

// Dry run check
if (args.Contains("--apply"))
{
    // Batch update
    var data = new List<ValueRange>();
    foreach (var (row, city) in updates)
    {
        data.Add(new ValueRange
        {
            Range = $"'{sheetName}'!H{row}",
            Values = [[city]]
        });
    }

    if (data.Count > 0)
    {
        var batchUpdate = new BatchUpdateValuesRequest
        {
            ValueInputOption = "RAW",
            Data = data
        };
        var result = await service.Spreadsheets.Values.BatchUpdate(batchUpdate, spreadsheetId).ExecuteAsync();
        Console.WriteLine($"\nUpdated {result.TotalUpdatedCells} cells!");
    }
}
else
{
    Console.WriteLine("\nDry run. Add --apply to actually write to the sheet.");
}

string? ExtractCity(string address)
{
    if (string.IsNullOrWhiteSpace(address))
        return null;

    // Split by comma
    var parts = address.Split(',').Select(p => p.Trim()).Where(p => p.Length > 0).ToList();

    if (parts.Count == 0)
        return null;

    var first = parts[0];

    // "Московская обл." or "Московская область" — city is the next part
    if (first.Contains("обл", StringComparison.OrdinalIgnoreCase))
    {
        if (parts.Count >= 2)
        {
            var candidate = parts[1].Trim();
            if (knownCities.Contains(candidate))
                return candidate;
            // Could be a city even if not in list
            if (candidate.Length >= 3 && candidate.Length <= 30 && !candidate.Contains("ул") && !candidate.Contains("пр"))
                return candidate;
        }
        return "Московская обл.";
    }

    // Direct city match
    if (knownCities.Contains(first))
        return first;

    // Single word, looks like a city name (starts with uppercase, no street markers)
    if (parts.Count == 1 && first.Length >= 3 && first.Length <= 30
        && char.IsUpper(first[0])
        && !first.Contains("ул") && !first.Contains("ш.") && !first.Contains("пр")
        && !first.Contains("ТРЦ") && !first.Contains("ТЦ")
        && !first.Contains("агент") && !first.Contains("работ"))
    {
        // Could be a city
        if (knownCities.Contains(first))
            return first;
    }

    // First part looks like a city (before comma, known)
    if (first.Length >= 2 && first.Length <= 30 && knownCities.Contains(first))
        return first;

    return null;
}
