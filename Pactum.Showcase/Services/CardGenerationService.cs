using Anthropic.SDK;
using Anthropic.SDK.Messaging;

namespace Pactum.Showcase.Services;

public class CardGenerationService
{
    private readonly AnthropicClient _client;
    private readonly string _model;
    private readonly DriveFileService _driveService;
    private readonly PromptService _promptService;
    private readonly ILogger<CardGenerationService> _logger;

    public CardGenerationService(IConfiguration config, DriveFileService driveService,
        PromptService promptService, ILogger<CardGenerationService> logger)
    {
        _driveService = driveService;
        _promptService = promptService;
        _logger = logger;

        var apiKey = config["Anthropic:ApiKey"]
            ?? throw new InvalidOperationException("Anthropic:ApiKey not configured");
        _client = new AnthropicClient(apiKey);
        _model = config["Anthropic:Model"] ?? "claude-sonnet-4-20250514";
    }

    public async Task<(bool success, string message)> GenerateCardAsync(string externalId)
    {
        _logger.LogInformation("Generating card for {Id}", externalId);

        // Get prompt template
        var promptTemplate = _promptService.GetDefault("card");
        if (promptTemplate == null)
            return (false, "Не найден промпт для генерации карточек");

        // Read schema
        var (schemaHtml, schemaError) = await _driveService.ReadMainDocAsync(externalId);
        if (schemaError != null && schemaHtml == null)
            return (false, $"Не удалось прочитать схему: {schemaError}");

        var schemaText = StripHtml(schemaHtml ?? "");
        if (string.IsNullOrWhiteSpace(schemaText))
            return (false, "Схема пуста");

        // Build prompt
        var prompt = promptTemplate.Text.Replace("{schema_text}", schemaText);

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
