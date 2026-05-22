using Anthropic.SDK;
using Anthropic.SDK.Messaging;

namespace Pactum.Showcase.Services;

public class ImagePromptService
{
    private readonly AnthropicClient _client;
    private readonly string _model;
    private readonly DriveFileService _driveService;
    private readonly PromptService _promptService;
    private readonly ILogger<ImagePromptService> _logger;

    public ImagePromptService(IConfiguration config, DriveFileService driveService,
        PromptService promptService, ILogger<ImagePromptService> logger)
    {
        _driveService = driveService;
        _promptService = promptService;
        _logger = logger;

        var apiKey = config["Anthropic:ApiKey"]
            ?? throw new InvalidOperationException("Anthropic:ApiKey not configured");
        _client = new AnthropicClient(apiKey);
        _model = config["Anthropic:Model"] ?? "claude-sonnet-4-20250514";
    }

    public async Task<(bool success, string text)> GeneratePromptAsync(string externalId, string? promptId = null)
    {
        await _promptService.EnsureLoadedAsync();
        var template = promptId != null
            ? _promptService.GetById(promptId)
            : _promptService.GetDefault("image-prompt");
        if (template == null)
            return (false, "Не найден шаблон промпта для картинок");

        var (cardHtml, cardError) = await _driveService.ReadCardDocAsync(externalId);
        if (cardError != null && cardHtml == null)
            return (false, $"Не удалось прочитать карточку: {cardError}");

        var cardText = StripHtml(cardHtml ?? "");
        if (string.IsNullOrWhiteSpace(cardText))
            return (false, "Карточка пуста. Сначала сгенерируйте карточку.");

        var prompt = template.Text.Replace("{card_text}", cardText);

        try
        {
            var message = await _client.Messages.GetClaudeMessageAsync(new MessageParameters
            {
                Model = _model,
                MaxTokens = 1024,
                Messages = [new Message(RoleType.User, prompt)]
            });

            var result = message.Content.FirstOrDefault()?.ToString()?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(result))
                return (false, "Claude вернул пустой ответ");

            // Strip surrounding quotes if any
            if (result.StartsWith('"') && result.EndsWith('"'))
                result = result[1..^1];

            _logger.LogInformation("Image prompt generated for {Id}, length: {Len}", externalId, result.Length);
            return (true, result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to generate image prompt for {Id}", externalId);
            return (false, $"Ошибка генерации: {ex.Message}");
        }
    }

    private static string StripHtml(string html)
    {
        var diagIdx = html.IndexOf("<hr/><small");
        if (diagIdx > 0)
            html = html[..diagIdx];

        return System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", " ")
            .Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">")
            .Replace("&quot;", "\"").Replace("&#39;", "'")
            .Replace("  ", " ").Trim();
    }
}
