using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace Json.Da
{
    class ApplicationContext: DbContext
    {
        public DbSet<Employee> Employees { get; set; }
        public DbSet<Discipline> Disciplines { get; set; }
        public DbSet<Syllabus> Syllabuses { get; set; }
         //public DbSet<Workload> Workload{ get; set; }

        public ApplicationContext()
        {
            //Database.EnsureDeleted();
            //Database.EnsureCreated();
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer("Server=(localdb)\\mssqllocaldb;Database=helloappdb;Trusted_Connection=True;");
        }
    }
}
