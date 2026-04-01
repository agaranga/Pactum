namespace Pactum.Showcase.Models;

public class Business
{
    public int Row { get; set; }
    public string Name { get; set; } = string.Empty;
    public Dictionary<string, string> Properties { get; set; } = new();
    public string? GeneratedDescription { get; set; }
    public bool IsGenerating { get; set; }
}
