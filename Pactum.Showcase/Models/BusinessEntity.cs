namespace Pactum.Showcase.Models;

public class BusinessEntity
{
    public int Id { get; set; }
    public int SheetRow { get; set; }
    public string ExternalId { get; set; } = "";  // ID from folder, e.g. "ID021539"
    public string Name { get; set; } = "";
    public string City { get; set; } = "";
    public string Address { get; set; } = "";
    public string Phone { get; set; } = "";
    public string Manager { get; set; } = "";
    public string Status { get; set; } = "";
    public string ActivityType { get; set; } = "";
    public string Description { get; set; } = "";
    public string Link { get; set; } = "";
    public string PropertiesJson { get; set; } = "{}"; // all columns as JSON
    public string? GeneratedDescription { get; set; }
    public DateTime LoadedAt { get; set; } = DateTime.UtcNow;
}
