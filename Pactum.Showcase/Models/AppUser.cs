namespace Pactum.Showcase.Models;

public class AppUser
{
    public string Id { get; set; } = "";
    public string Username { get; set; } = "";
    public string PasswordHash { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string Role { get; set; } = "Operator"; // Admin, Operator
}
