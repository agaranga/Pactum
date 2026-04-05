using System.Security.Claims;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.EntityFrameworkCore;
using Pactum.Showcase.Components;
using Pactum.Showcase.Data;
using Pactum.Showcase.Middleware;
using Pactum.Showcase.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {
        options.LoginPath = "/login";
        options.LogoutPath = "/logout";
        options.ExpireTimeSpan = TimeSpan.FromDays(30);
        options.SlidingExpiration = true;
    });
builder.Services.AddAuthorization();
builder.Services.AddCascadingAuthenticationState();

// Database
var dbPath = Path.Combine(builder.Environment.ContentRootPath, "pactum.db");
builder.Services.AddDbContextFactory<AppDbContext>(options =>
    options.UseSqlite($"Data Source={dbPath}"));

// Services
builder.Services.AddSingleton<GoogleSheetsApiService>();
builder.Services.AddSingleton<DataService>();
builder.Services.AddSingleton<DriveFileService>();
builder.Services.AddSingleton<DescriptionService>();
builder.Services.AddSingleton<GoogleOAuthService>();
builder.Services.AddSingleton<CardGenerationService>();
builder.Services.AddSingleton<IUserService, ConfigUserService>();

var app = builder.Build();

// Ensure DB created
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    await db.Database.EnsureCreatedAsync();
}

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
}

app.UseMiddleware<IpWhitelistMiddleware>();
app.UseStatusCodePagesWithReExecute("/not-found", createScopeForStatusCodePages: true);
app.UseAuthentication();
app.UseAuthorization();
app.UseAntiforgery();

app.MapPost("/api/auth/login", async (HttpContext ctx, IUserService userService) =>
{
    var form = await ctx.Request.ReadFormAsync();
    var username = form["username"].ToString();
    var password = form["password"].ToString();

    var user = await userService.ValidateAsync(username, password);
    if (user == null)
    {
        ctx.Response.Redirect("/login?error=1");
        return;
    }

    var claims = new List<Claim>
    {
        new(ClaimTypes.NameIdentifier, user.Id),
        new(ClaimTypes.Name, user.DisplayName),
        new(ClaimTypes.Role, user.Role)
    };

    var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
    await ctx.SignInAsync(
        CookieAuthenticationDefaults.AuthenticationScheme,
        new ClaimsPrincipal(identity),
        new AuthenticationProperties { IsPersistent = true, ExpiresUtc = DateTimeOffset.UtcNow.AddDays(30) });

    ctx.Response.Redirect("/");
}).DisableAntiforgery();

app.MapGet("/api/auth/logout", async (HttpContext ctx) =>
{
    await ctx.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
    ctx.Response.Redirect("/login");
});

app.MapGet("/api/auth/google", (HttpContext ctx, GoogleOAuthService oauth) =>
{
    var scheme = ctx.Request.Scheme;
    var host = ctx.Request.Host;
    var redirectUri = $"{scheme}://{host}/api/auth/google-callback";
    var url = oauth.GetAuthorizationUrl(redirectUri);
    ctx.Response.Redirect(url);
}).RequireAuthorization();

app.MapGet("/api/auth/google-callback", async (HttpContext ctx, GoogleOAuthService oauth) =>
{
    var code = ctx.Request.Query["code"].ToString();
    if (string.IsNullOrWhiteSpace(code))
    {
        ctx.Response.Redirect("/admin?google=error");
        return;
    }

    var scheme = ctx.Request.Scheme;
    var host = ctx.Request.Host;
    var redirectUri = $"{scheme}://{host}/api/auth/google-callback";

    var success = await oauth.ExchangeCodeAsync(code, redirectUri);
    ctx.Response.Redirect(success ? "/admin?google=ok" : "/admin?google=error");
});

app.MapGet("/api/images/{fileId}", async (string fileId, DriveFileService driveService) =>
{
    var result = await driveService.GetImageAsync(fileId);
    if (result == null)
        return Results.NotFound();
    return Results.File(result.Value.data, result.Value.mimeType);
}).RequireAuthorization();

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
