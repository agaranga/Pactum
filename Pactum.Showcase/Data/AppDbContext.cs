using Microsoft.EntityFrameworkCore;
using Pactum.Showcase.Models;

namespace Pactum.Showcase.Data;

public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

    public DbSet<BusinessEntity> Businesses => Set<BusinessEntity>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<BusinessEntity>(e =>
        {
            e.HasKey(b => b.Id);
            e.HasIndex(b => b.ExternalId);
        });
    }
}
