using Microsoft.EntityFrameworkCore;
using TechReportToDB.Data.Entities;

namespace TechReportToDB.Data
{
    internal class AppDbContext : DbContext
    {
        public DbSet<Job> Jobs { get; set; }
        public DbSet<Tool> Tools { get; set; }
        public DbSet<Kit> Kits { get; set; }
        public DbSet<KitTool> KitTools { get; set; }
        public DbSet<DD> DDs { get; set; }
        public DbSet<MWD> MWDs { get; set; }
        public DbSet<Construction> Constructions { get; set; }

        public AppDbContext(DbContextOptions<AppDbContext> opt) : base(opt)
        {
            Database.EnsureCreated();
        }
    }
}
