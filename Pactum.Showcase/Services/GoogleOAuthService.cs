using System.Text.Json;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Drive.v3;
using Google.Apis.Docs.v1;
using Google.Apis.Services;

namespace Pactum.Showcase.Services;

public class GoogleOAuthService
{
    private readonly string _clientId;
    private readonly string _clientSecret;
    private readonly string? _railwayApiToken;
    private readonly string? _railwayProjectId;
    private readonly string? _railwayServiceId;
    private readonly string? _railwayEnvironmentId;
    private readonly ILogger<GoogleOAuthService> _logger;
    private UserCredential? _credential;
    private readonly SemaphoreSlim _lock = new(1, 1);

    public GoogleOAuthService(IConfiguration config, ILogger<GoogleOAuthService> logger)
    {
        _logger = logger;
        _clientId = config["GoogleOAuth:ClientId"]
            ?? throw new InvalidOperationException("GoogleOAuth:ClientId not configured");
        _clientSecret = config["GoogleOAuth:ClientSecret"]
            ?? throw new InvalidOperationException("GoogleOAuth:ClientSecret not configured");

        _railwayApiToken = config["Railway:ApiToken"];
        _railwayProjectId = config["Railway:ProjectId"];
        _railwayServiceId = config["Railway:ServiceId"];
        _railwayEnvironmentId = config["Railway:EnvironmentId"];

        // Try restore token from environment variable
        var savedToken = config["GoogleOAuth:RefreshToken"];
        if (!string.IsNullOrWhiteSpace(savedToken))
        {
            try
            {
                var flow = CreateFlow();
                var token = new TokenResponse
                {
                    RefreshToken = savedToken,
                    IssuedUtc = DateTime.UtcNow
                };
                _credential = new UserCredential(flow, "user", token);
                _logger.LogInformation("Google OAuth token restored from environment variable");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to restore OAuth token from env");
            }
        }
    }

    public bool IsAuthorized => _credential?.Token?.RefreshToken != null;

    public string GetDiagnostics() =>
        $"Railway API Token: {(_railwayApiToken != null ? "set" : "NOT SET")}, " +
        $"ProjectId: {(_railwayProjectId != null ? "set" : "NOT SET")}, " +
        $"ServiceId: {(_railwayServiceId != null ? "set" : "NOT SET")}, " +
        $"EnvironmentId: {(_railwayEnvironmentId != null ? "set" : "NOT SET")}, " +
        $"HasRefreshToken: {_credential?.Token?.RefreshToken != null}";

    public string GetAuthorizationUrl(string redirectUri)
    {
        var flow = CreateFlow();
        var uri = flow.CreateAuthorizationCodeRequest(redirectUri).Build();
        var uriStr = uri.ToString();
        // Force consent to always get refresh token
        if (!uriStr.Contains("prompt="))
            uriStr += "&prompt=consent";
        else
            uriStr = uriStr.Replace("prompt=", "prompt=consent");
        return uriStr;
    }

    public async Task<bool> ExchangeCodeAsync(string code, string redirectUri)
    {
        await _lock.WaitAsync();
        try
        {
            var flow = CreateFlow();
            var token = await flow.ExchangeCodeForTokenAsync("user", code, redirectUri, CancellationToken.None);
            _credential = new UserCredential(flow, "user", token);

            _logger.LogInformation("Google OAuth token obtained. RefreshToken present: {HasRefresh}",
                !string.IsNullOrWhiteSpace(token.RefreshToken));

            // Save refresh token to Railway variable
            if (!string.IsNullOrWhiteSpace(token.RefreshToken))
                await SaveRefreshTokenToRailwayAsync(token.RefreshToken);
            else
                _logger.LogWarning("No refresh token received from Google. Token may already exist — try revoking access at https://myaccount.google.com/permissions and reconnecting");

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to exchange OAuth code");
            return false;
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task<DriveService?> GetDriveServiceAsync()
    {
        var cred = await GetCredentialAsync();
        if (cred == null) return null;
        return new DriveService(new BaseClientService.Initializer { HttpClientInitializer = cred });
    }

    public async Task<DocsService?> GetDocsServiceAsync()
    {
        var cred = await GetCredentialAsync();
        if (cred == null) return null;
        return new DocsService(new BaseClientService.Initializer { HttpClientInitializer = cred });
    }

    private async Task<UserCredential?> GetCredentialAsync()
    {
        if (_credential == null) return null;

        if (_credential.Token.IsStale)
        {
            try
            {
                await _credential.RefreshTokenAsync(CancellationToken.None);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to refresh OAuth token");
                return null;
            }
        }

        return _credential;
    }

    private GoogleAuthorizationCodeFlow CreateFlow()
    {
        return new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
        {
            ClientSecrets = new ClientSecrets
            {
                ClientId = _clientId,
                ClientSecret = _clientSecret
            },
            Scopes = [DriveService.Scope.DriveFile, DocsService.Scope.Documents]
        });
    }

    private async Task SaveRefreshTokenToRailwayAsync(string refreshToken)
    {
        if (string.IsNullOrWhiteSpace(_railwayApiToken)
            || string.IsNullOrWhiteSpace(_railwayProjectId)
            || string.IsNullOrWhiteSpace(_railwayServiceId)
            || string.IsNullOrWhiteSpace(_railwayEnvironmentId))
        {
            _logger.LogWarning("Railway API config not set, cannot persist OAuth token");
            return;
        }

        try
        {
            using var http = new HttpClient();
            http.DefaultRequestHeaders.Add("Authorization", $"Bearer {_railwayApiToken}");

            var query = new
            {
                query = @"mutation($input: VariableUpsertInput!) { variableUpsert(input: $input) }",
                variables = new
                {
                    input = new
                    {
                        projectId = _railwayProjectId,
                        serviceId = _railwayServiceId,
                        environmentId = _railwayEnvironmentId,
                        name = "GoogleOAuth__RefreshToken",
                        value = refreshToken
                    }
                }
            };

            var response = await http.PostAsJsonAsync("https://backboard.railway.app/graphql/v2", query);
            var body = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode)
                _logger.LogInformation("OAuth refresh token saved to Railway variable");
            else
                _logger.LogWarning("Failed to save token to Railway: {Status} {Body}", response.StatusCode, body);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to save OAuth token to Railway");
        }
    }
}
