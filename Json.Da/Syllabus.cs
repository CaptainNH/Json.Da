using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Json.Da
{
    class Syllabus
    {
        public Discipline Predmet { get; set; }

        public int Year { get; set; }

        public string Direction { get; set; }

        public string Semester { get; set; }

        public int Hours { get; set; }
    }
}
