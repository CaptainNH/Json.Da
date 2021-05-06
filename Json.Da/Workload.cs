using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Json.Da
{
    class Workload
    {
        public int Id { get; set; }

        public int Year { get; set; }
        public Employee Employee { get; set; }

        public Syllabus Plan { get; set; }
    }
}
