using System.Text.Json;
using System.Text.Json.Serialization;

namespace Pactum.Showcase.Services;

public class ImageGenerationService
{
    private readonly HttpClient _http;
    private readonly DriveFileService _drive;
    private readonly string _apiKey;
    private readonly string _loraUrl;
    private readonly double _loraScale;
    private readonly int _numImages;
    private readonly int _steps;
    private readonly double _guidance;
    private readonly string _imageSize;
    private readonly ILogger<ImageGenerationService> _logger;

    public ImageGenerationService(IConfiguration config, DriveFileService drive,
        IHttpClientFactory httpFactory, ILogger<ImageGenerationService> logger)
    {
        _http = httpFactory.CreateClient();
        _http.Timeout = TimeSpan.FromMinutes(5);
        _drive = drive;
        _logger = logger;

        _apiKey = config["FalAi:ApiKey"]
            ?? throw new InvalidOperationException("FalAi:ApiKey not configured");
        _loraUrl = config["FalAi:LoraUrl"]
            ?? throw new InvalidOperationException("FalAi:LoraUrl not configured");
        _loraScale = config.GetValue("FalAi:LoraScale", 0.8);
        _numImages = config.GetValue("FalAi:NumImages", 4);
        _steps = config.GetValue("FalAi:Steps", 28);
        _guidance = config.GetValue("FalAi:GuidanceScale", 3.5);
        _imageSize = config["FalAi:ImageSize"] ?? "square_hd";
    }

    public async Task<(bool success, string message, int saved)> GenerateAsync(string externalId, string prompt)
    {
        if (string.IsNullOrWhiteSpace(prompt))
            return (false, "Промпт пустой", 0);

        _logger.LogInformation("Generating images for {Id}, prompt length: {Len}", externalId, prompt.Length);

        // 1. Call fal.ai
        var request = new HttpRequestMessage(HttpMethod.Post, "https://fal.run/fal-ai/flux-lora");
        request.Headers.Add("Authorization", $"Key {_apiKey}");
        request.Content = JsonContent.Create(new FalRequest
        {
            Prompt = prompt,
            Loras = [new FalLora { Path = _loraUrl, Scale = _loraScale }],
            ImageSize = _imageSize,
            NumImages = _numImages,
            GuidanceScale = _guidance,
            NumInferenceSteps = _steps
        });

        FalResponse? falResponse;
        try
        {
            var response = await _http.SendAsync(request);
            var body = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning("fal.ai error {Status}: {Body}", response.StatusCode, body);
                return (false, $"fal.ai вернул {(int)response.StatusCode}: {body}", 0);
            }

            falResponse = JsonSerializer.Deserialize<FalResponse>(body, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "fal.ai request failed");
            return (false, $"Ошибка запроса к fal.ai: {ex.Message}", 0);
        }

        if (falResponse?.Images == null || falResponse.Images.Length == 0)
            return (false, "fal.ai не вернул картинки", 0);

        _logger.LogInformation("fal.ai returned {Count} images for {Id}", falResponse.Images.Length, externalId);

        // 2. Get next available index in Drive folder
        int startIdx = await _drive.GetNextGenIndexAsync(externalId);

        // 3. Download and upload each
        int saved = 0;
        for (int i = 0; i < falResponse.Images.Length; i++)
        {
            try
            {
                var imageUrl = falResponse.Images[i].Url;
                if (string.IsNullOrWhiteSpace(imageUrl)) continue;

                var imageBytes = await _http.GetByteArrayAsync(imageUrl);
                var fileName = $"{externalId}_GEN_{startIdx + i}.png";

                var (ok, error) = await _drive.UploadImageAsync(externalId, fileName, imageBytes, "image/png");
                if (ok)
                    saved++;
                else
                    _logger.LogWarning("Failed to save image {File}: {Error}", fileName, error);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to process image {Idx}", i);
            }
        }

        return (saved > 0, $"Сгенерировано и сохранено {saved} картинок", saved);
    }

    private class FalRequest
    {
        [JsonPropertyName("prompt")]
        public string Prompt { get; set; } = "";

        [JsonPropertyName("loras")]
        public FalLora[] Loras { get; set; } = [];

        [JsonPropertyName("image_size")]
        public string ImageSize { get; set; } = "square_hd";

        [JsonPropertyName("num_images")]
        public int NumImages { get; set; } = 4;

        [JsonPropertyName("guidance_scale")]
        public double GuidanceScale { get; set; } = 3.5;

        [JsonPropertyName("num_inference_steps")]
        public int NumInferenceSteps { get; set; } = 28;
    }

    private class FalLora
    {
        [JsonPropertyName("path")]
        public string Path { get; set; } = "";

        [JsonPropertyName("scale")]
        public double Scale { get; set; } = 0.8;
    }

    private class FalResponse
    {
        [JsonPropertyName("images")]
        public FalImage[]? Images { get; set; }
    }

    private class FalImage
    {
        [JsonPropertyName("url")]
        public string? Url { get; set; }
    }
}
