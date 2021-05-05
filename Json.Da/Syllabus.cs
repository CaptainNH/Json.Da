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

        public string SubjectName { get; set; }
        public Discipline Predmet { get; set; }
        public int Year { get; set; }
        public string Direction { get; set; }//
        public string Profile { get; set; }//
        public int Semester { get; set; }//
        public int CreditUnits { get; set; }//
        public string Hours { get; set; }//
        public string CourseWork { get; set; }//
        public string SumIndependentWork { get; set; }//
        public string InteractiveWatch { get; set; }//
        public bool Test { get; set; }//
        public bool Exam { get; set; }//
        public int Lectures { get; set; }//
        public int LaboratoryExercises { get; set; }//
        public int Workshops { get; set; }//
        public string TypesOfLessons { get; set; }//
        public double AuditoryLessons { get; set; }//

        public Syllabus()
        {
            Year = 0;
            Direction = "";
            Profile = "";
            Semester = 0;
            CreditUnits = 0;
            Hours = "";
            CourseWork = "-";
            SumIndependentWork = "";
            Lectures = 0;
            InteractiveWatch = "";
            LaboratoryExercises = 0;
            Workshops = 0;
            TypesOfLessons = "";
            AuditoryLessons = 0;
        }

        public void SetYear(IXLWorksheet workSheet, string cellName)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(cellName).Value.ToString()))
                this.Year = Convert.ToInt32(workSheet.Cell(cellName).Value.ToString());
        }

        public void SetDirectionAndProfile(IXLWorksheet workSheet, string cellName)
        {
            string[] separators = new string[] {"Направленность программы", "Направление подготовки", "Профиль",
                "Профиль:", "Профили", "Направление", "Программа"}; //разделители направления и профиля
            var directionAndProfile = workSheet.Cell(cellName).Value.ToString().Split(separators, StringSplitOptions.RemoveEmptyEntries); //Сплит по разделителям
            this.Direction = directionAndProfile[0].Trim(' ', ',', ':'); //Получить направление
            this.Profile = "";
            if (directionAndProfile.Length > 1)
                this.Profile = "Профиль: " + directionAndProfile[1].Trim(' ', ':'); //Получить профиль, если он есть            
        }

        public void SetCreditUnits(IXLWorksheet workSheet, int row)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(row,8).Value.ToString()))
                this.CreditUnits = Convert.ToInt32(workSheet.Cell(row, 8).Value.ToString().Trim(' '));
        }

        public void SetHours(IXLWorksheet workSheet, int row)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(row, 11).Value.ToString()))
                this.Hours = workSheet.Cell(row,11).Value.ToString().Trim(' ') + " час.";
        }

        public void SetAuditoryLessons(IXLWorksheet workSheet, int row)
        {
            double h = 0, iw = 0;
            if (!string.IsNullOrEmpty(workSheet.Cell(row, 11).Value.ToString()))
                h = Convert.ToDouble(workSheet.Cell(row, 11).Value.ToString());
            if (!string.IsNullOrEmpty(workSheet.Cell(row, 14).Value.ToString()))
                iw = Convert.ToDouble(workSheet.Cell(row, 14).Value.ToString());
            this.AuditoryLessons = h - iw;
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

        public void SetTypesOfLessons()
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
            this.TypesOfLessons = s;
        }

        public void SetSemester(IXLWorksheet workSheet, int column)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(2, column).Value.ToString()))
                this.Semester = Convert.ToInt32(workSheet.Cell(2, column).Value.ToString().Split()[1]);
        }

        public void SetSubjectName(IXLWorksheet workSheet, int row)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(row, 3).Value.ToString()))
                this.SubjectName = workSheet.Cell(row, 3).Value.ToString();
        }
    }
}
