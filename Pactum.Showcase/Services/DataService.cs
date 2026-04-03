using System.Text.Json;
using Microsoft.EntityFrameworkCore;
using Pactum.Showcase.Data;
using Pactum.Showcase.Models;

namespace Pactum.Showcase.Services;

public class DataService
{
    private readonly IDbContextFactory<AppDbContext> _dbFactory;
    private readonly GoogleSheetsApiService _sheetsApi;
    private readonly ILogger<DataService> _logger;

    private List<BusinessEntity>? _cache;
    private readonly SemaphoreSlim _lock = new(1, 1);

    public DataService(IDbContextFactory<AppDbContext> dbFactory, GoogleSheetsApiService sheetsApi, ILogger<DataService> logger)
    {
        _dbFactory = dbFactory;
        _sheetsApi = sheetsApi;
        _logger = logger;
    }

    public async Task<List<BusinessEntity>> GetAllAsync()
    {
        if (_cache != null)
            return _cache;

        await _lock.WaitAsync();
        try
        {
            if (_cache != null)
                return _cache;

            using var db = await _dbFactory.CreateDbContextAsync();
            var data = await db.Businesses.OrderBy(b => b.SheetRow).ToListAsync();

            if (data.Count == 0)
            {
                _logger.LogInformation("Database empty, auto-loading from Google Sheets");
                await ReloadFromSheetsInternalAsync();
                using var db2 = await _dbFactory.CreateDbContextAsync();
                data = await db2.Businesses.OrderBy(b => b.SheetRow).ToListAsync();
            }

            _cache = data;
            return _cache;
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task<BusinessEntity?> GetByIdAsync(int id)
    {
        var all = await GetAllAsync();
        return all.FirstOrDefault(b => b.Id == id);
    }

    public async Task<int> GetCountAsync()
    {
        var all = await GetAllAsync();
        return all.Count;
    }

    public async Task ClearAsync()
    {
        await _lock.WaitAsync();
        try
        {
            using var db = await _dbFactory.CreateDbContextAsync();
            db.Businesses.RemoveRange(db.Businesses);
            await db.SaveChangesAsync();
            _cache = [];
            _logger.LogInformation("Database cleared");
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task<int> ReloadFromSheetsAsync()
    {
        await _lock.WaitAsync();
        try
        {
            var count = await ReloadFromSheetsInternalAsync();
            using var db = await _dbFactory.CreateDbContextAsync();
            _cache = await db.Businesses.OrderBy(b => b.SheetRow).ToListAsync();
            return count;
        }
        finally
        {
            _lock.Release();
        }
    }

    private async Task<int> ReloadFromSheetsInternalAsync()
    {
        _logger.LogInformation("Reloading data from Google Sheets");

        var allRows = await _sheetsApi.ReadAllDataAsync();
        var headerRow = 2;
        if (allRows.Count < headerRow + 1)
        {
            _logger.LogWarning("Sheet has too few rows: {Count}", allRows.Count);
            return 0;
        }

        var headers = allRows[headerRow - 1]
            .Select(h => h?.ToString()?.Trim() ?? "")
            .ToList();

        int Col(string name) => headers.FindIndex(h =>
            h.Equals(name, StringComparison.OrdinalIgnoreCase));

        var colName = Col("Название");
        var colCity = Col("Город");
        var colAddr = Col("Адрес");
        var colPhone = Col("Мобильный");
        var colManager = Col("Менеджер");
        var colStatus = Col("Статус обработки");
        var colActivity = Col("Вид деятельности");
        var colDesc = Col("Краткое описание");
        var colLink = Col("ССЫЛКА");
        var colId = Col("id");

        var entities = new List<BusinessEntity>();

        for (int i = headerRow; i < allRows.Count; i++)
        {
            var row = allRows[i];
            string Get(int col) => col >= 0 && col < row.Count
                ? row[col]?.ToString()?.Trim() ?? "" : "";

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

            var name = Get(colName);
            if (string.IsNullOrWhiteSpace(name))
                name = $"Бизнес #{i + 1}";

            entities.Add(new BusinessEntity
            {
                SheetRow = i + 1,
                ExternalId = Get(colId),
                Name = name,
                City = Get(colCity),
                Address = Get(colAddr),
                Phone = Get(colPhone),
                Manager = Get(colManager),
                Status = Get(colStatus),
                ActivityType = Get(colActivity),
                Description = Get(colDesc),
                Link = Get(colLink),
                PropertiesJson = JsonSerializer.Serialize(properties),
                LoadedAt = DateTime.UtcNow
            });
        }

        using var db = await _dbFactory.CreateDbContextAsync();
        db.Businesses.RemoveRange(db.Businesses);
        await db.Businesses.AddRangeAsync(entities);
        await db.SaveChangesAsync();

        _logger.LogInformation("Loaded {Count} businesses into database", entities.Count);
        return entities.Count;
    }
}
