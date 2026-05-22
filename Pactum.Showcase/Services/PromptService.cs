using System.Text.Json;
using Pactum.Showcase.Models;

namespace Pactum.Showcase.Services;

public class PromptService
{
    private const string ConfigFileName = "prompts.json";
    private readonly DriveFileService _drive;
    private readonly ILogger<PromptService> _logger;
    private List<PromptTemplate> _prompts = [];
    private bool _loaded;
    private readonly SemaphoreSlim _lock = new(1, 1);

    public PromptService(DriveFileService drive, ILogger<PromptService> logger)
    {
        _drive = drive;
        _logger = logger;
        _prompts = CreateDefaults();
    }

    public async Task EnsureLoadedAsync()
    {
        if (_loaded) return;

        await _lock.WaitAsync();
        try
        {
            if (_loaded) return;
            try
            {
                var json = await _drive.ReadConfigFileAsync(ConfigFileName);
                if (!string.IsNullOrWhiteSpace(json))
                {
                    var loaded = JsonSerializer.Deserialize<List<PromptTemplate>>(json,
                        new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                    if (loaded != null && loaded.Count > 0)
                    {
                        _prompts = loaded;
                        _logger.LogInformation("Loaded {Count} prompts from Drive", _prompts.Count);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to load prompts from Drive, using defaults");
            }
            _loaded = true;
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task<List<PromptTemplate>> GetAllAsync()
    {
        await EnsureLoadedAsync();
        return _prompts.ToList();
    }

    public async Task<PromptTemplate?> GetByIdAsync(string id)
    {
        await EnsureLoadedAsync();
        return _prompts.FirstOrDefault(p => p.Id == id);
    }

    public async Task<PromptTemplate?> GetDefaultAsync(string category)
    {
        await EnsureLoadedAsync();
        return _prompts.FirstOrDefault(p => p.Category == category && p.IsDefault)
            ?? _prompts.FirstOrDefault(p => p.Category == category);
    }

    public List<PromptTemplate> GetAll() => _prompts.ToList();
    public PromptTemplate? GetById(string id) => _prompts.FirstOrDefault(p => p.Id == id);
    public PromptTemplate? GetDefault(string category) =>
        _prompts.FirstOrDefault(p => p.Category == category && p.IsDefault)
        ?? _prompts.FirstOrDefault(p => p.Category == category);

    public async Task<(bool success, string? error)> UpdateAsync(string id, string name, string text)
    {
        await EnsureLoadedAsync();
        var p = _prompts.FirstOrDefault(x => x.Id == id);
        if (p == null) return (false, "Промпт не найден");
        p.Name = name;
        p.Text = text;
        return await SaveAsync();
    }

    public async Task<(bool success, string? error)> CreateAsync(string name, string category, string text)
    {
        await EnsureLoadedAsync();
        var p = new PromptTemplate
        {
            Id = $"{category}-{Guid.NewGuid().ToString("N")[..8]}",
            Name = name,
            Category = category,
            Text = text,
            IsDefault = false
        };
        _prompts.Add(p);
        return await SaveAsync();
    }

    public async Task<(bool success, string? error)> DuplicateAsync(string id)
    {
        await EnsureLoadedAsync();
        var src = _prompts.FirstOrDefault(p => p.Id == id);
        if (src == null) return (false, "Промпт не найден");
        var copy = new PromptTemplate
        {
            Id = $"{src.Category}-{Guid.NewGuid().ToString("N")[..8]}",
            Name = $"{src.Name} (копия)",
            Category = src.Category,
            Text = src.Text,
            IsDefault = false
        };
        _prompts.Add(copy);
        return await SaveAsync();
    }

    public async Task<(bool success, string? error)> DeleteAsync(string id)
    {
        await EnsureLoadedAsync();
        var p = _prompts.FirstOrDefault(x => x.Id == id);
        if (p == null) return (false, "Промпт не найден");
        if (p.IsDefault) return (false, "Нельзя удалить промпт по умолчанию. Сначала назначьте другой по умолчанию.");
        _prompts.Remove(p);
        return await SaveAsync();
    }

    public async Task<(bool success, string? error)> SetDefaultAsync(string id)
    {
        await EnsureLoadedAsync();
        var p = _prompts.FirstOrDefault(x => x.Id == id);
        if (p == null) return (false, "Промпт не найден");
        // Reset default for all in same category, set for this one
        foreach (var other in _prompts.Where(x => x.Category == p.Category))
            other.IsDefault = false;
        p.IsDefault = true;
        return await SaveAsync();
    }

    private async Task<(bool success, string? error)> SaveAsync()
    {
        var json = JsonSerializer.Serialize(_prompts, new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        });
        return await _drive.SaveConfigFileAsync(ConfigFileName, json);
    }

    private static List<PromptTemplate> CreateDefaults()
    {
        return
        [
            new PromptTemplate
            {
                Id = "card-default",
                Name = "Карточка бизнеса (основной)",
                Category = "card",
                IsDefault = true,
                Text = """
                    Ты — профессиональный копирайтер, специализирующийся на продаже готового бизнеса. Твоя задача — на основе схемы (исходных данных о бизнесе) написать продающую карточку объявления.

                    ФОРМАТ КАРТОЧКИ (строго соблюдай структуру и порядок разделов):

                    1. **Заголовок:** Одна строка. Формат: "Ключевое название бизнеса | Главные преимущества через запятую, цифры, площадь, уникальные факты"

                    2. **Текст объявления:**

                       Начни с жирного абзаца: "Продается действующий/ее [тип бизнеса] в [локация]."
                       Затем 2-3 предложения — суть бизнеса, сколько лет работает, ключевые факты. НЕ упоминай причину продажи в начале — она будет в конце.

                       **Почему это выгодное приобретение:**
                       Перечисли 3-5 ключевых преимуществ в формате:
                       ✅ **[Преимущество жирным].** Развернутое пояснение на 2-3 предложения. Конкретика, цифры, факты. Не общие слова.

                       **Что входит в актив:**
                       Перечисли по категориям в формате:
                       - **Категория:** описание.
                       Категории могут включать: Оборудование, Помещение, Интеллектуальная собственность, Каналы продаж, Клиентская база, Команда, Юридическая структура — используй только те, для которых есть данные в схеме.

                       **Финансовые показатели:**
                       Жирным: выручка, прибыль, расходы — только если есть в исходных данных. Если данных нет — напиши "предоставляются по запросу".

                       **Для кого это предложение:**
                       Каждое предложение начинай с новой строки.
                       Начни с "Это вариант для предпринимателя, который хочет...".
                       Подчеркни, что всё готово, не нужно начинать с нуля.
                       Всего 3-4 предложения, каждое на отдельной строке.

                       **Стоимость бизнеса:** [сумма] (если есть данные по окупаемости — добавь).

                       **Причина продажи:** 1 предложение из исходных данных. Этот раздел должен быть только здесь, в конце текста. Не дублируй причину продажи в начале карточки или в других разделах.

                       **📞 Заинтересовало предложение?** Звоните, чтобы обсудить детали, [2-3 конкретных действия: запросить отчеты, посмотреть помещение, встретиться с управленцем — в зависимости от специфики бизнеса].

                    ПРАВИЛА:
                    - Пиши на русском языке.
                    - Не выдумывай данные. Используй только то, что есть в схеме.
                    - Если данных по какому-то пункту нет — пропусти его или напиши "по запросу".
                    - Стиль: деловой, но живой. Не сухой отчет, а продающий текст. Без восклицательных знаков в каждом предложении.
                    - Объем текста объявления: 2000-4000 символов.
                    - Все цифры, названия, адреса бери строго из схемы.
                    - Используй ✅ только в разделе "Почему это выгодное приобретение". Во всех остальных списках используй тире (-).
                    - Причина продажи упоминается ТОЛЬКО один раз, в разделе "Причина продажи" в конце текста.

                    СХЕМА БИЗНЕСА:
                    {schema_text}
                    """
            },
            new PromptTemplate
            {
                Id = "image-prompt-default",
                Name = "Промпт для картинки (isocutaway)",
                Category = "image-prompt",
                IsDefault = true,
                Text = """
                    Ты составляешь промпт для генерации изометрической cutaway-иллюстрации бизнеса через FLUX-LoRA. Стиль уже зашит в LoRA с триггером isocutaway (изометрия 30°, cutaway-вид, пастельная палитра, cute 3D-render diorama, мягкое освещение). В промпт стиль писать НЕ надо — описывай только содержимое.

                    СТРОГАЯ СТРУКТУРА (одна строка через запятые):
                    isocutaway, isometric cutaway diorama of a [тип бизнеса], [зона 1 с конкретными предметами], [зона 2 с конкретными предметами], ..., [зона N], [сотрудники в одинаковой униформе], [действия людей], [опционально: окружение здания]

                    ПРАВИЛА:
                    1. Триггер isocutaway всегда первым словом, без дублирования.
                    2. НЕ описывай стиль (никаких "miniature 3D render", "pastel palette", "cute diorama style", "isometric angle", "highly detailed", "8k", "masterpiece").
                    3. НЕ упоминай камеру, текстуру, освещение, цвета подробно.
                    4. НЕ добавляй негативные промпты.
                    5. 6-10 функциональных зон бизнеса. Меньше — пусто, больше — модель путается.
                    6. В каждой зоне 2-4 конкретных предмета (не "стеллажи", а "metal shelving with spare parts and oil cans").
                    7. Сотрудники в одинаковой униформе единого цвета — это маркер бизнеса.
                    8. От 2 до 4 человек в кадре. В конце промпта ОБЯЗАТЕЛЬНО добавь две буквальные фразы: "2 to 4 people in total" и "clean detailed faces, no artifacts, no extra people in background". Обе фразы должны присутствовать в финальном промпте.
                    9. Окружение здания (asphalt road, brick wall, sidewalk) — опционально.
                    10. Английский язык. 60-120 слов. Одна строка через запятые. Никаких bullet-списков, переносов, markdown.
                    11. Если есть риск что модель добавит надписи — допиши в конец: no text, no labels, no signs.

                    На входе ты получаешь карточку продажи бизнеса (на русском). На основе её содержимого составь промпт строго по структуре выше.

                    Верни ТОЛЬКО готовый промпт одной строкой. Без пояснений, без префиксов вроде "Промпт:", без кавычек.

                    КАРТОЧКА БИЗНЕСА:
                    {card_text}
                    """
            }
        ];
    }
}
