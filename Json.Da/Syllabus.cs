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

        public int Year { get; set; }
        /*{
            get { return Year; }
            set 
            { 
                if (value>1900)
                {
                    Year = value;
                }
            } 
        }*/
        public string Direction { get; set; }//
        public string Profile { get; set; }//
        public string StudyProgram { get; set; }//
        bool IsGraduateSchool { get; set; }//
        public string Standart {get; set;}//
        public string Protocol { get; set; }//
        public string EdForm { get; set; }//
        public string DirectionAbbreviation { get; set; }//
        public string Director { get; set; }//
        public string Position { get; set; }

        public string SubjectName { get; set; }
        public Discipline Predmet { get; set; }
        public string Semester { get; set; }//
        public string Course { get; set; }//
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
            Semester = "-";
            IsGraduateSchool = false;
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
        public static int OnlyYear(IXLWorksheet workSheet, string cellName)
        {
            int a = 0;  
            if (!string.IsNullOrEmpty(workSheet.Cell(cellName).Value.ToString()))
            {
                a = Convert.ToInt32(workSheet.Cell(cellName).Value.ToString());
            }
            return a;
        }
        public static string SetOnlyDirection(IXLWorksheet workSheet, string cellName)
        {
            string[] separators = new string[] {"Направленность программы", "Направление подготовки", "Профиль",
                "Профиль:", "Профили", "Направление", "Программа"}; //разделители направления и профиля
            var directionAndProfile = workSheet.Cell(cellName).Value.ToString().Split(separators, StringSplitOptions.RemoveEmptyEntries); //Сплит по разделителям
            string direct = directionAndProfile[0].Trim(' ', ',', ':');
            return direct; //Получить направление

        }
        public void SetDirectionAndProfile(IXLWorksheet workSheet, string cellName)
        {
            string[] separators = new string[] {"Направленность программы", "Направление подготовки", "Профиль",
                "Профиль:", "Профили", "Направление", "Программа"}; //разделители направления и профиля
            var directionAndProfile = workSheet.Cell(cellName).Value.ToString().Split(separators, StringSplitOptions.RemoveEmptyEntries); //Сплит по разделителям
            this.Direction = directionAndProfile[0].Trim(' ', ',', ':'); //Получить направление
            this.Profile = "";
            if (directionAndProfile.Length > 1)
                this.Profile = directionAndProfile[1].Trim(' ', ':', '\"'); //Получить профиль, если он есть            
        }

        public void SetStudyProgram(IXLWorksheet workSheet, string cellName)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(cellName).Value.ToString()))
                this.StudyProgram = workSheet.Cell(cellName).Value.ToString().Replace("  ", " ").Trim(' ').Split()[2];
            if (this.StudyProgram == "аспирантуры")
                this.IsGraduateSchool = true;
        }


        public void SetStandart(IXLWorksheet workSheet, string cellName)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(cellName).Value.ToString()))
            {
                var s = workSheet.Cell(cellName).Value.ToString().Split(new string[] { "от" }, StringSplitOptions.RemoveEmptyEntries);
                this.Standart = s[1].Trim(' ') + " г. " + s[0].Trim(' ');
            }
                
        }

        public void SetProtocol(IXLWorksheet workSheet, string cellName)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(cellName).Value.ToString()))
            {
                var s = workSheet.Cell(cellName).Value.ToString().Split(new string[] { "Протокол", "от" }, StringSplitOptions.RemoveEmptyEntries);
                this.Protocol = s[1].Trim(' ') + " г. " + s[0].Trim(' ');
            }

        }

        public void SetEdForm(IXLWorksheet workSheet, string cellName, string cellNameAspir)
        {
            var s = new string[2];
            if (IsGraduateSchool)
                s = workSheet.Cell(cellNameAspir).Value.ToString().Split(':');
            else
                s = workSheet.Cell(cellName).Value.ToString().Split(':');
            this.EdForm = s[1].Trim(' ') + " " + s[0].ToLower();
        }

        public void SetDirectionAbbreviation(IXLWorksheet workSheet, string cellName)
        {
            //Создаем аббревиатуры направлений.
            string directionName = workSheet.Cell(cellName).Value.ToString();
            string abbreviation = "";
            if (this.StudyProgram == "магистратуры")
                abbreviation = "МАГИ_";
            else if (this.StudyProgram == "аспирантуры")
            {
                abbreviation = "АСПИР_";
                if (this.Profile.Contains("логика"))
                    abbreviation += "МЛ";
                else if (this.Profile.Contains("уравнения"))
                    abbreviation += "ДУ";
                this.DirectionAbbreviation = abbreviation;
            }
            if (directionName.Contains("  "))
                directionName = directionName.Replace("  ", " ");
            string[] splittedDirectionName = directionName.Split(' ');
            if (splittedDirectionName.Contains("Прикладная"))
                abbreviation += "ПМ";
            else if (splittedDirectionName.Contains("Педагогическое"))
                abbreviation += "ПОМИ";
            else if (splittedDirectionName.Contains("Информатика"))
                abbreviation += "ИВТ";
            else
                abbreviation += "МАТ";
            this.DirectionAbbreviation = abbreviation;
        }

        public void SetDirestor(string dir1, string dir2, string dir3)
        {
            string s = dir1;
            if (this.StudyProgram == "аспирантуры")
            {
                s = dir2;
            }
            else if (StudyProgram == "магистратуры")
            {
                s = dir3;
            }
            this.Director = s;
        }

        public void SetPosition(string pos1, string pos2, string pos3)
        {
            string s = pos1;
            if (this.StudyProgram == "аспирантуры")
            {
                s = pos2;
            }
            else if (StudyProgram == "магистратуры")
            {
                s = pos3;
            }
            this.Position = s;
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
            this.Test = tests.Contains(this.Semester);
        }

        public void SetExam(IXLWorksheet workSheet, int row)
        {
            string exam = workSheet.Cell(row, 4).Value.ToString();
            this.Exam = exam.Contains(this.Semester);
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
            if (this.IsGraduateSchool)
            {
                if (!string.IsNullOrEmpty(workSheet.Cell(1, column).Value.ToString()))
                    this.Semester = workSheet.Cell(1, column).Value.ToString().Split()[1];
            }
            else
            {
                if (!string.IsNullOrEmpty(workSheet.Cell(2, column).Value.ToString()))
                    this.Semester = workSheet.Cell(2, column).Value.ToString().Split()[1];
            } 
        }

        public void SetCourse()
        {
            int semester;
            if (this.Semester == "A")
                semester = 10;
            else
                semester = Convert.ToInt32(this.Semester);
            if (this.IsGraduateSchool)
            {
                this.Course = this.Semester;
                this.Semester = "-";
            }
            else
                this.Course = ((semester + 1) / 2).ToString();
        }

        public void SetSubjectName(IXLWorksheet workSheet, int row)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(row, 3).Value.ToString()))
                this.SubjectName = workSheet.Cell(row, 3).Value.ToString();
        }
    }
}
