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

        public AppContext()
        {
            Database.EnsureCreated();
        }
    }
}
