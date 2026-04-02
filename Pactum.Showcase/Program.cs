using System.Security.Claims;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Pactum.Showcase.Components;
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

builder.Services.AddSingleton<GoogleSheetsApiService>();
builder.Services.AddSingleton<GoogleSheetsService>();
builder.Services.AddSingleton<DescriptionService>();
builder.Services.AddSingleton<IUserService, ConfigUserService>();

var app = builder.Build();

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

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
