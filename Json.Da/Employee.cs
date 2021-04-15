using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Json.Da
{
    class Employee
    {
        public string Surname { get; set; }

        public string Name { get; set; }

        public string Fathername { get; set; }

        public string Position { get; set; }

        public string Rank { get; set; }

        public double Rate { get; set; }

        public string Chair { get; set; }
    }
}
