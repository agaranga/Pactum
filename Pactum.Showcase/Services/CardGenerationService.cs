using Anthropic.SDK;
using Anthropic.SDK.Messaging;

namespace Pactum.Showcase.Services;

public class CardGenerationService
{
    private readonly AnthropicClient _client;
    private readonly string _model;
    private readonly DriveFileService _driveService;
    private readonly ILogger<CardGenerationService> _logger;

    private const string PromptTemplate = """
        Ты — профессиональный копирайтер, специализирующийся на продаже готового бизнеса. Твоя задача — на основе схемы (исходных данных о бизнесе) написать продающую карточку объявления.

        ФОРМАТ КАРТОЧКИ (строго соблюдай структуру и порядок разделов):

        1. **Заголовок:** Одна строка. Формат: "Ключевое название бизнеса | Главные преимущества через запятую, цифры, площадь, уникальные факты"

        2. **Текст объявления:**

           Начни с жирного абзаца: "Продается действующий/ее [тип бизнеса] в [локация]."
           Затем 2-3 предложения — суть бизнеса, сколько лет работает, почему продается. Без воды, только факты.

           **Почему это выгодное приобретение:**
           Перечисли 3-5 ключевых преимуществ в формате:
           ✅ **[Преимущество жирным].** Развернутое пояснение на 2-3 предложения. Конкретика, цифры, факты. Не общие слова.

           **Что входит в актив:**
           Перечисли по категориям в формате:
           🔹 **Категория:** описание.
           Категории могут включать: Оборудование, Помещение, Интеллектуальная собственность, Каналы продаж, Клиентская база, Команда, Юридическая структура — используй только те, для которых есть данные в схеме.

           **Финансовые показатели:**
           Жирным: выручка, прибыль, расходы — только если есть в исходных данных. Если данных нет — напиши "предоставляются по запросу".

           **Для кого это предложение:** 3-4 предложения. Начни с "Это вариант для предпринимателя, который хочет...". Подчеркни, что всё готово, не нужно начинать с нуля.

           **Стоимость бизнеса:** [сумма] (если есть данные по окупаемости — добавь).

           **Причина продажи:** 1 предложение из исходных данных.

           **📞 Заинтересовало предложение?** Звоните, чтобы обсудить детали, [2-3 конкретных действия: запросить отчеты, посмотреть помещение, встретиться с управленцем — в зависимости от специфики бизнеса].

        ПРАВИЛА:
        - Пиши на русском языке.
        - Не выдумывай данные. Используй только то, что есть в схеме.
        - Если данных по какому-то пункту нет — пропусти его или напиши "по запросу".
        - Стиль: деловой, но живой. Не сухой отчет, а продающий текст. Без восклицательных знаков в каждом предложении.
        - Объем текста объявления: 2000-4000 символов.
        - Все цифры, названия, адреса бери строго из схемы.

        СХЕМА БИЗНЕСА:
        {schema_text}
        """;

    public CardGenerationService(IConfiguration config, DriveFileService driveService, ILogger<CardGenerationService> logger)
    {
        _driveService = driveService;
        _logger = logger;

        var apiKey = config["Anthropic:ApiKey"]
            ?? throw new InvalidOperationException("Anthropic:ApiKey not configured");
        _client = new AnthropicClient(apiKey);
        _model = config["Anthropic:Model"] ?? "claude-sonnet-4-20250514";
    }

    /// <summary>
    /// Generate card for a single business: read schema, call Claude, save docx.
    /// </summary>
    public async Task<(bool success, string message)> GenerateCardAsync(string externalId)
    {
        _logger.LogInformation("Generating card for {Id}", externalId);

        // 1. Read schema (main doc)
        var (schemaHtml, schemaError) = await _driveService.ReadMainDocAsync(externalId);
        if (schemaError != null && schemaHtml == null)
            return (false, $"Не удалось прочитать схему: {schemaError}");

        // Strip HTML tags to get plain text for the prompt
        var schemaText = StripHtml(schemaHtml ?? "");
        if (string.IsNullOrWhiteSpace(schemaText))
            return (false, "Схема пуста");

        // 2. Call Claude API
        var prompt = PromptTemplate.Replace("{schema_text}", schemaText);

        try
        {
            var message = await _client.Messages.GetClaudeMessageAsync(new MessageParameters
            {
                Model = _model,
                MaxTokens = 4096,
                Messages = [new Message(RoleType.User, prompt)]
            });

            var cardText = message.Content.FirstOrDefault()?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(cardText))
                return (false, "Claude вернул пустой ответ");

            _logger.LogInformation("Card generated for {Id}, length: {Len}", externalId, cardText.Length);

            // 3. Save as docx to Drive
            var saveResult = await _driveService.SaveCardDocxAsync(externalId, cardText);
            if (!saveResult.success)
                return (false, $"Карточка сгенерирована, но не удалось сохранить: {saveResult.error}");

            return (true, "Карточка сгенерирована и сохранена");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to generate card for {Id}", externalId);
            return (false, $"Ошибка генерации: {ex.Message}");
        }
    }

    private static string StripHtml(string html)
    {
        // Remove diagnostic section
        var diagIdx = html.IndexOf("<hr/><small");
        if (diagIdx > 0)
            html = html[..diagIdx];

        return System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", " ")
            .Replace("&amp;", "&")
            .Replace("&lt;", "<")
            .Replace("&gt;", ">")
            .Replace("&quot;", "\"")
            .Replace("&#39;", "'")
            .Replace("  ", " ")
            .Trim();
    }
}
