using ContractorsWorkAPI.Model;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using System.ComponentModel.DataAnnotations.Schema;

namespace ContractorsWorkAPI.Configuration
{
    public class FilesConfiguration  : IEntityTypeConfiguration<Files>
    {
        public void Configure(EntityTypeBuilder<Files> builder)
        {
            builder.HasKey(x => x.Id);
            builder.Property(x => x.Id).HasColumnName("id");
            builder.Property(x => x.Name).HasColumnName("name").IsRequired();
            builder.Property(x => x.Path).HasColumnName("path").IsRequired();
            builder.Property(x => x.CreateDate).HasColumnName("create_date").IsRequired();
        }
    }
}
