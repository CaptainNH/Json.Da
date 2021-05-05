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
        public string CourseWork { get; set; } 
        public string SumIndependentWork { get; set; }
        public string InteractiveWatch { get; set; }
        public bool Test { get; set; }
        public bool Exam { get; set; }
        public int Lectures { get; set; }
        public int LaboratoryExercises { get; set; }
        public int Workshops { get; set; }
        public string TypesOfLessons { get; set; }

        public Syllabus()
        {
            Direction = "";
            Profile = "";
            CreditUnits = 0;
            StudyHours = "";
            SumIndependentWork = "";
            CourseWork = "-";
            Lectures = 0;
            LaboratoryExercises = 0;
            Workshops = 0;
            TypesOfLessons = "";
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

        public void SetCreditUnits(IXLWorksheet workSheet, int row)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(row,8).Value.ToString()))
                this.CreditUnits = Convert.ToInt32(workSheet.Cell(8, row).Value.ToString().Trim(' '));
        }

        public void SetStudyHours(IXLWorksheet workSheet, int row)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(row, 11).Value.ToString()))
                this.StudyHours = workSheet.Cell(row,11).Value.ToString().Trim(' ') + " час.";
        }

        public void SetSumIndependentWork(IXLWorksheet workSheet, int row)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell( row,14).Value.ToString()))
                this.SumIndependentWork = workSheet.Cell(row,14).Value.ToString().Trim(' ');
        }

        public void SetInteractiveWatch(IXLWorksheet workSheet, int row)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(row,16).Value.ToString()))
                this.InteractiveWatch = workSheet.Cell(row,16).Value.ToString().Trim(' ');
        }

        public void SetCourseWork(IXLWorksheet workSheet, int row)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(row, 7).Value.ToString()))
                this.CourseWork = workSheet.Cell(row, 7).Value.ToString().Trim(' ');
        }

        public void SetTests(IXLWorksheet workSheet, int row)
        {
            string GradedTest = workSheet.Cell(row,6).Value.ToString();
            string test = workSheet.Cell(row, 5).Value.ToString();
            string tests = GradedTest + test;
            this.Test = tests.Contains(this.Semester.ToString());
        }

        public void SetExam(IXLWorksheet workSheet, int row)
        {
            string exam = workSheet.Cell(row, 4).Value.ToString();
            this.Exam = exam.Contains(this.Semester.ToString());
        }

        public void SetLestures(IXLWorksheet workSheet, int row, int column)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(row, column+2).Value.ToString()))
                this.Lectures = Convert.ToInt32(workSheet.Cell(row, column+2).Value.ToString().Trim(' '));
        }

        public void SetLaboratoryExercises(IXLWorksheet workSheet, int row, int column)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(row, column + 3).Value.ToString()))
                this.LaboratoryExercises = Convert.ToInt32(workSheet.Cell(row, column+3).Value.ToString().Trim(' '));
        }

        public void SetWorkshops(IXLWorksheet workSheet, int row, int column)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(row, column + 4).Value.ToString()))
                this.Workshops = Convert.ToInt32(workSheet.Cell(row, column+4).Value.ToString().Trim(' '));
        }

        string CreateTypesOfLessons()
        {
            string s = "";
            var list = new List<string>();
            if (Lectures != 0)
                list.Add("лекционных");
            if (Workshops != 0)
                list.Add("практических");
            if (LaboratoryExercises != 0)
                list.Add("лабораторных");
            if (list.Count == 1)
                s = list[0];
            else if (list.Count == 2)
                s = list[0] + " и " + list[1];
            else if (list.Count == 3)
                s = list[0] + ", " + list[1] + " и " + list[2];
            return s;
        }

        /*public void Set(IXLWorksheet workSheet, int index)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(, index).Value.ToString()))
                this. = workSheet.Cell(, index).Value.ToString().Trim(' ');
        }*/

    }
}
