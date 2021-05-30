using System;
using System.Collections.Generic;
using System.Linq;
//BANANA
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.ComponentModel.DataAnnotations.Schema;

namespace Json.Da
{
    class Discipline
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public string Competencies { get; set; }
        public string Compiler { get; set; }
        public string NamePat { get; set; }
        public string Date { get; set; }
        public string Koi { get; set; }
        public string DisciplineTarget { get; set; }
        public string OPOP { get; set; }
        public string Know { get; set; }
        public string BeAbleTo { get; set; }
        public string Own { get; set; }
        public string ControlTasks { get; set; }
        public string TestTasks { get; set; }
        public string QuestionForTest { get; set; }
        public string InformationSupportOfDiscipline { get; set; }
        public string LogisticsOfTheDiscipline { get; set; }
        public string UpdateSheet { get; set; }
        [NotMapped]
        public List<string> EducTechn { get; set; }
        [NotMapped]
        public List<string> DiscMap { get; set; }
        [NotMapped]
        public string MethodologyAssessment { get; set; }
    }
}
