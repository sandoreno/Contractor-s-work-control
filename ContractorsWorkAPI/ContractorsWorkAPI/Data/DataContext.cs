using ContractorsWorkAPI.Configuration;
using ContractorsWorkAPI.Model;
using Microsoft.EntityFrameworkCore;

namespace ContractorsWorkAPI.Data
{
    public class DataContext : DbContext
    {
        public DataContext(DbContextOptions<DataContext> options)
          : base(options) { }

        public DbSet<Files>? Files { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.ApplyConfiguration(new FilesConfiguration());
        }


    }
}
