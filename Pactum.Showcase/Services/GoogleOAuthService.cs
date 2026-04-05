using System.Text.Json;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Drive.v3;
using Google.Apis.Docs.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace Pactum.Showcase.Services;

/// <summary>
/// Manages OAuth2 token for writing to Google Drive on behalf of a real user.
/// Service account is used for reading; OAuth is used for creating/writing files.
/// </summary>
public class GoogleOAuthService
{
    private readonly string _clientId;
    private readonly string _clientSecret;
    private readonly string _tokenPath;
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
        _tokenPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Pactum", "google-oauth-token");
    }

    public bool IsAuthorized => _credential?.Token?.AccessToken != null;

    public string GetAuthorizationUrl(string redirectUri)
    {
        var flow = CreateFlow();
        return flow.CreateAuthorizationCodeRequest(redirectUri).Build().ToString();
    }

    public async Task<bool> ExchangeCodeAsync(string code, string redirectUri)
    {
        await _lock.WaitAsync();
        try
        {
            var flow = CreateFlow();
            var token = await flow.ExchangeCodeForTokenAsync("user", code, redirectUri, CancellationToken.None);
            _credential = new UserCredential(flow, "user", token);

            // Save token
            Directory.CreateDirectory(Path.GetDirectoryName(_tokenPath)!);
            var json = JsonSerializer.Serialize(new
            {
                token.AccessToken,
                token.RefreshToken,
                token.ExpiresInSeconds,
                Issued = token.IssuedUtc
            });
            await File.WriteAllTextAsync(_tokenPath + ".json", json);

            _logger.LogInformation("Google OAuth token saved");
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
        if (_credential != null)
        {
            if (_credential.Token.IsStale)
                await _credential.RefreshTokenAsync(CancellationToken.None);
            return _credential;
        }

        // Try load saved token
        var tokenFile = _tokenPath + ".json";
        if (File.Exists(tokenFile))
        {
            try
            {
                var json = await File.ReadAllTextAsync(tokenFile);
                var saved = JsonSerializer.Deserialize<SavedToken>(json);
                if (saved?.RefreshToken != null)
                {
                    var flow = CreateFlow();
                    var token = new TokenResponse
                    {
                        AccessToken = saved.AccessToken,
                        RefreshToken = saved.RefreshToken,
                        ExpiresInSeconds = saved.ExpiresInSeconds,
                        IssuedUtc = saved.Issued
                    };
                    _credential = new UserCredential(flow, "user", token);
                    if (_credential.Token.IsStale)
                        await _credential.RefreshTokenAsync(CancellationToken.None);
                    return _credential;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to load saved OAuth token");
            }
        }

        return null;
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

    private class SavedToken
    {
        public string? AccessToken { get; set; }
        public string? RefreshToken { get; set; }
        public long? ExpiresInSeconds { get; set; }
        public DateTime Issued { get; set; }
    }
}
