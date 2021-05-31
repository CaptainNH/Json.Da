using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Json.Da
{
    class EducTechn
    {
        public int Id { get; set; }
        public string Theme { get; set; }
        public string ActivityType { get; set; }
        public int NumberOfHours { get; set; }
        public string ActiveForms { get; set; }
        public string InteractiveForms { get; set; }
    }
    class DisccMap
    {
        public int Id { get; set; }
        public string DiscQuestion { get; set; }
        public int Lection { get; set; }
        public int Practice { get; set; }
        public string Content { get; set; }
        public int Hours { get; set; }
        public string FormsControl { get; set; }
        public int Min { get; set; }
        public int Max { get; set; }
        public string Literature { get; set; }
    }
    class ResultMark
    {
        public int Id { get; set; }
        public string FormsControl { get; set; }
        public string CurrentControl { get; set; }
        public string First { get; set; }
        public string Second { get; set; }
        public string Third { get; set; }
        public string Fourth { get; set; }
    }
}
