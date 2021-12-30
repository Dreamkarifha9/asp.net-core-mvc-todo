using AspnetCoreTODO.Models;
using Microsoft.EntityFrameworkCore;

namespace AspnetCoreTODO.Data
{
    public class ApplicationDbContext : DbContext
    {
        public virtual DbSet<Todo> Todos { get; set; }
    public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
        : base(options) {

        }
        protected override void OnModelCreating(ModelBuilder modelBuilder) {
      base.OnModelCreating(modelBuilder);
    }
    }
}