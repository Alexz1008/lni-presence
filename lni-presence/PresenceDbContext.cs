using Microsoft.EntityFrameworkCore;

namespace lni_presence;

public class PresenceDbContext : DbContext
{
    public PresenceDbContext(DbContextOptions<PresenceDbContext> options) : base(options) { }

    public DbSet<PresenceRecord> PresenceLog => Set<PresenceRecord>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<PresenceRecord>(entity =>
        {
            entity.ToTable("PresenceLog");
            entity.HasIndex(e => new { e.UserId, e.IsCurrent });
        });
    }
}
