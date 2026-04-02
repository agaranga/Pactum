using Microsoft.AspNetCore.Identity;
using Pactum.Showcase.Models;

namespace Pactum.Showcase.Services;

public class ConfigUserService : IUserService
{
    private readonly List<AppUser> _users;
    private readonly PasswordHasher<AppUser> _hasher = new();

    public ConfigUserService(IConfiguration config)
    {
        _users = config.GetSection("Users").Get<List<AppUser>>() ?? [];
    }

    public Task<AppUser?> ValidateAsync(string username, string password)
    {
        var user = _users.FirstOrDefault(u =>
            u.Username.Equals(username, StringComparison.OrdinalIgnoreCase));

        if (user == null)
            return Task.FromResult<AppUser?>(null);

        var result = _hasher.VerifyHashedPassword(user, user.PasswordHash, password);
        return Task.FromResult(result == PasswordVerificationResult.Failed ? null : user);
    }

    public Task<AppUser?> GetByIdAsync(string id)
    {
        return Task.FromResult(_users.FirstOrDefault(u => u.Id == id));
    }

    public Task<List<AppUser>> GetAllAsync()
    {
        return Task.FromResult(_users.ToList());
    }
}
