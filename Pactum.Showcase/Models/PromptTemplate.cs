namespace Pactum.Showcase.Models;

public class PromptTemplate
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Category { get; set; } = ""; // "card", "image", etc.
    public string Text { get; set; } = "";
    public bool IsDefault { get; set; }
}
