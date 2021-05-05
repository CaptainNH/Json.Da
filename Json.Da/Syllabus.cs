using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Json.Da
{
    class Syllabus
    {
        public int Id { get; set; }

        public Discipline Predmet { get; set; }

        public int Year { get; set; }

        public string Direction { get; set; }

        public string Profile { get; set; }

        public int Semester { get; set; }

        public string SubjectName { get; set; }

        public int Hours { get; set; }

        public int CreditUnits { get; set; }

        public string StudyHours { get; set; }

        //public string a { get; set; }
        public string SumIndependentWork { get; set; }

        public string InteractiveWatch { get; set; }
        public Syllabus()
        {
            Direction = "none";
            Profile = "none";
            CreditUnits = 0;
            StudyHours = "none";
            SumIndependentWork = "";
        }

        public void SetDirectionAndProfile(IXLWorksheet workSheet)
        {
            string[] separators = new string[] {"Направленность программы", "Направление подготовки", "Профиль",
                "Профиль:", "Профили", "Направление", "Программа"}; //разделители направления и профиля
            var directionAndProfile = workSheet.Cell("B18").Value.ToString().Split(separators, StringSplitOptions.RemoveEmptyEntries); //Сплит по разделителям
            this.Direction = directionAndProfile[0].Trim(' ', ',', ':'); //Получить направление
            this.Profile = "";
            if (directionAndProfile.Length > 1)
                this.Profile = "Профиль: " + directionAndProfile[1].Trim(' ', ':'); //Получить профиль, если он есть            
        }

        public void SetCreditUnits(IXLWorksheet workSheet, int index)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(8, index).Value.ToString()))
                this.CreditUnits = Convert.ToInt32(workSheet.Cell(8, index).Value.ToString().Trim(' '));
        }

        public void SetStudyHours(IXLWorksheet workSheet, int index)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(11, index).Value.ToString()))
                this.StudyHours = workSheet.Cell(11, index).Value.ToString().Trim(' ') + " час.";
        }

        public void SetSumIndependentWork(IXLWorksheet workSheet, int index)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(14, index).Value.ToString()))
                this.SumIndependentWork = workSheet.Cell(14, index).Value.ToString().Trim(' ');
        }

        public void SetInteractiveWatch(IXLWorksheet workSheet, int index)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(16, index).Value.ToString()))
                this.InteractiveWatch = workSheet.Cell(16, index).Value.ToString().Trim(' ');
        }

        public void Set(IXLWorksheet workSheet, int index)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(, index).Value.ToString()))
                this. = workSheet.Cell(, index).Value.ToString().Trim(' ');
        }

        /*public void Set(IXLWorksheet workSheet, int index)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(, index).Value.ToString()))
                this. = workSheet.Cell(, index).Value.ToString().Trim(' ');
        }*/

    }
}
