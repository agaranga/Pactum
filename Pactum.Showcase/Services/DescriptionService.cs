using Anthropic.SDK;
using Anthropic.SDK.Messaging;
using Pactum.Showcase.Models;

namespace Pactum.Showcase.Services;

public class DescriptionService
{
    private readonly AnthropicClient _client;
    private readonly string _model;
    private readonly string _promptTemplate;
    private readonly ILogger<DescriptionService> _logger;

    public DescriptionService(IConfiguration config, ILogger<DescriptionService> logger)
    {
        _logger = logger;

        var apiKey = config["Anthropic:ApiKey"]
            ?? throw new InvalidOperationException("Anthropic:ApiKey not configured");

        _client = new AnthropicClient(apiKey);
        _model = config["Anthropic:Model"] ?? "claude-sonnet-4-20250514";
        _promptTemplate = config["Anthropic:PromptTemplate"]
            ?? """
               Ты — профессиональный копирайтер. Составь красивое, структурированное описание бизнеса
               на основе следующих параметров. Описание должно быть на русском языке, привлекательным
               для потенциальных партнёров и инвесторов. Используй markdown-форматирование.

               Параметры бизнеса:
               {PROPERTIES}
               """;
    }

    public async Task<string> GenerateDescriptionAsync(Business business)
    {
        var propertiesText = string.Join("\n",
            business.Properties.Select(kv => $"- **{kv.Key}**: {kv.Value}"));

        var prompt = _promptTemplate.Replace("{PROPERTIES}", propertiesText);

        _logger.LogInformation("Generating description for: {Name}", business.Name);

        var message = await _client.Messages.GetClaudeMessageAsync(new MessageParameters
        {
            Model = _model,
            MaxTokens = 2048,
            Messages = [new Message(RoleType.User, prompt)]
        });

        var result = message.Content.FirstOrDefault()?.ToString() ?? "Не удалось сгенерировать описание.";

        _logger.LogInformation("Description generated for: {Name}, length: {Len}", business.Name, result.Length);
        return result;
    }
}
