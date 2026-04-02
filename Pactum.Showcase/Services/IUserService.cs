using Pactum.Showcase.Models;

namespace Pactum.Showcase.Services;

public interface IUserService
{
    Task<AppUser?> ValidateAsync(string username, string password);
    Task<AppUser?> GetByIdAsync(string id);
    Task<List<AppUser>> GetAllAsync();
}
