using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;

var credential = GoogleCredential.FromFile("Pactum.Showcase/google-credentials.json")
    .CreateScoped(SheetsService.Scope.Spreadsheets);
var sheets = new SheetsService(new BaseClientService.Initializer { HttpClientInitializer = credential });
var sid = "15prE_Se2EnRuukWUq7tutPm3S4yKvtirMJvQzp_VFhs";

// Only confident matches (both name+addr or unique brand)
var updates = new (int row, string id8, string desc)[]
{
    (41, "ID021593", "Пиццерия → Pizza Express 24, name+addr"),
    (95, "ID021533", "Цветы и кафе на Ленинском, name:4/4+addr"),
    (5,  "ID021540", "Пункт выдачи заказов, name+addr"),
    (6,  "ID021541", "Империал/Imperial, brand+addr"),
    (9,  "ID021545", "Пивной бар, name+addr"),
    (11, "ID021548", "Пивной паб на Измайловском, name+addr"),
    (14, "ID021553", "Мебельный магазин ТРЦ Бутово Молл, name+addr"),
    (66, "ID021559", "Хабанеро перец, unique name+addr"),
    (46, "ID021585", "КОНЮШНЯ/конный спорт, name+addr"),
    (29, "ID021572", "Ювелирная мастерская Каширское ш., name+addr"),
    (31, "ID021596", "Защитные плёнки, specific name"),
    (16, "ID021560", "Швейное производство Мытищи, name+addr"),
    (65, "ID021558", "ПВЗ OZON Коммунарка, addr:5/6"),
    (77, "ID021551", "Цветочный магазин Потаповская Роща, name+addr"),
    (51, "ID021582", "Косметология Большая Полянка, addr:3/4"),
    (10, "ID021546", "Кофейня Большая Семёновская, addr:3/3"),
    (81, "ID021550", "Кафе Волоколамское шоссе, addr:3/3"),
    (15, "ID021555", "Пекарня Варшавское шоссе, addr:3/3"),
    (64, "ID021567", "Древесный уголь Дмитров, addr:4/8"),
    (12, "ID021549", "Итальянское бистро, addr match"),
    (40, "ID021592", "Детский центр/KidsLOFT, name+addr"),
};

Console.WriteLine($"Applying {updates.Length} confident matches:\n");
foreach (var u in updates)
    Console.WriteLine($"  Row {u.row} = {u.id8} ({u.desc})");

var data = updates.Select(u => new ValueRange
{
    Range = $"'АнкетаMAX'!B{u.row}",
    Values = [[u.id8]]
}).ToList();

var batch = new BatchUpdateValuesRequest { ValueInputOption = "RAW", Data = data };
var result = await sheets.Spreadsheets.Values.BatchUpdate(batch, sid).ExecuteAsync();
Console.WriteLine($"\nUpdated {result.TotalUpdatedCells} cells!");
