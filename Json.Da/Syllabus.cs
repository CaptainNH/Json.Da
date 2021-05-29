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
        private bool IsGraduateSchool { get; set; }//
        public string Standart {get; set;}//
        public string Protocol { get; set; }//
        public string EdForm { get; set; }//
        public string DirectionAbbreviation { get; set; }//
        public string Director { get; set; }//
        public string Position { get; set; }

        public string SubjectName { get; set; }//
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
        public bool Consulting { get; set; }//
        public int Lectures { get; set; }//
        public int LaboratoryExercises { get; set; }//
        public int Workshops { get; set; }//
        public double IndependentWorkBySemester { get; set; }//
        public string TypesOfLessons { get; set; }//
        public double AuditoryLessons { get; set; }//
        public string SubjectIndex { get; set; }//
        public string SubjectIndexDecoding { get; set; }//
        public string Competencies { get; set; }//

        //SubjectCompetencies////
        //SubjectIndex////
        //DecodeSubgectIndex////
        //IndependentWorkBySemester////
        //Consulting////

        //sumLectures
        //sumLabs
        //sumWorkshops
        public Syllabus()
        {
            Semester = "-";
        }

        public Syllabus(List<Discipline> predmetlist, List<Syllabus> listSyllabus, IXLWorksheet workSheetTitle, IXLWorksheet workSheetPlan, IXLWorksheet workSheetComp, Dictionary<string, string> compDic, int row, int column)
        {


            SetYear(workSheetTitle, "T29");
            SetDirectionAndProfile(workSheetTitle, "B18");
            SetStudyProgram(workSheetTitle, "F14");
            SetStandart(workSheetTitle, "T31");
            SetProtocol(workSheetTitle, "A13");
            SetEdForm(workSheetTitle, "A31", "A30");
            SetDirectionAbbreviation(workSheetTitle, "B18");
            SetDirestor("А.М. Дигурова", "Б.В. Туаева", "Л.А. Агузарова");
            SetPosition("Проректор по УР", "Проректор по научной деятельности", "Первый проректор");

            SetSemester(workSheetPlan, column);
            SetAuditoryLessons(workSheetPlan, row);
            SetCourseWork(workSheetPlan, row);
            SetCreditUnits(workSheetPlan, row);
            SetExam(workSheetPlan, row);
            SetHours(workSheetPlan, row);
            SetSubjectName(workSheetPlan, row);
            SetInteractiveWatch(workSheetPlan, row);
            SetLaboratoryExercises(workSheetPlan, row, column);
            SetLestures(workSheetPlan, row, column);
            SetSumIndependentWork(workSheetPlan, row);
            SetTests(workSheetPlan, row);
            SetWorkshops(workSheetPlan, row, column);
            SetCourse();
            SetTypesOfLessons();
            SetCompetencies(workSheetPlan, row, compDic);
            SetSubjectIndex(workSheetPlan, row);
            DecodeSubjectIndex(workSheetPlan, row);
            SetIndependentWorkBySemester(workSheetPlan, row, column);
            SetConsulting(workSheetPlan, row);
            Predmet = predmetlist.Find(item => item.Name == this.SubjectName);
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
            this.CreditUnits = 0;
            if (!string.IsNullOrEmpty(workSheet.Cell(row, 8).Value.ToString()))
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
            this.SumIndependentWork = "";
            if (!string.IsNullOrEmpty(workSheet.Cell( row,14).Value.ToString()))
                this.SumIndependentWork = workSheet.Cell(row,14).Value.ToString().Trim(' ');
        }

        public void SetInteractiveWatch(IXLWorksheet workSheet, int row)
        {
            this.InteractiveWatch = "";
            if (!string.IsNullOrEmpty(workSheet.Cell(row,16).Value.ToString()))
                this.InteractiveWatch = workSheet.Cell(row,16).Value.ToString().Trim(' ');
        }

        public void SetCourseWork(IXLWorksheet workSheet, int row)
        {
            this.CourseWork = "-";
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
            this.Exam = false;
            string exam = workSheet.Cell(row, 4).Value.ToString();
            this.Exam = exam.Contains(this.Semester);
        }

        public void SetConsulting(IXLWorksheet workSheet, int row)
        {
            this.Consulting = false;
            string consulting = workSheet.Cell(row, 4).Value.ToString();
            this.Consulting = consulting.Contains(this.Semester);
        }

        public void SetLestures(IXLWorksheet workSheet, int row, int column)
        {
            this.Lectures = 0;
            if (!string.IsNullOrEmpty(workSheet.Cell(row, column+2).Value.ToString()))
                this.Lectures = Convert.ToInt32(workSheet.Cell(row, column+2).Value.ToString().Trim(' '));
        }

        public void SetLaboratoryExercises(IXLWorksheet workSheet, int row, int column)
        {
            this.LaboratoryExercises = 0;
            if (!string.IsNullOrEmpty(workSheet.Cell(row, column + 3).Value.ToString()))
                this.LaboratoryExercises = Convert.ToInt32(workSheet.Cell(row, column+3).Value.ToString().Trim(' '));
        }

        public void SetWorkshops(IXLWorksheet workSheet, int row, int column)
        {
            this.Workshops = 0;
            if (!string.IsNullOrEmpty(workSheet.Cell(row, column + 4).Value.ToString()))
                this.Workshops = Convert.ToInt32(workSheet.Cell(row, column+4).Value.ToString().Trim(' '));
        }

        public void SetIndependentWorkBySemester(IXLWorksheet workSheet, int row, int column)
        {
            this.IndependentWorkBySemester = 0;
            if (!string.IsNullOrEmpty(workSheet.Cell(row, column + 5).Value.ToString()))
                this.IndependentWorkBySemester = Convert.ToDouble(workSheet.Cell(row, column + 5).Value.ToString().Trim(' '));
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
                this.Semester = "-";
                if (!string.IsNullOrEmpty(workSheet.Cell(1, column).Value.ToString()))
                    this.Semester = workSheet.Cell(1, column).Value.ToString().Split()[1];
            }
            else
            {
                this.Semester = "-";
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
            this.SubjectName = "?";
            if (!string.IsNullOrEmpty(workSheet.Cell(row, 3).Value.ToString()))
                this.SubjectName = workSheet.Cell(row, 3).Value.ToString();
        }

        public void SetSubjectIndex(IXLWorksheet workSheet, int row)
        {
            this.SubjectIndex = "?";
            if (!string.IsNullOrEmpty(workSheet.Cell(row, 2).Value.ToString()))
                this.SubjectIndex= workSheet.Cell(row, 2).Value.ToString();
        }

        public void DecodeSubjectIndex(IXLWorksheet workSheet, int row)
        {
            string subsectionName = "";
            string blockName = "";
            string[] s = this.SubjectIndex.Split('.');
            string subjectIndexDecoding = "";
            int i = row;
            while (!string.IsNullOrEmpty(workSheet.Cell(i,2).Value.ToString()))
                i--;
            subsectionName = workSheet.Cell(i, 1).Value.ToString().Trim(' ');

            while (workSheet.Cell(i, 1).Value.ToString().Split()[0].ToLower() != "блок")
                i--;
            string[] ss = workSheet.Cell(i, 1).Value.ToString().Trim(' ').Split('.');
            blockName = ss[0] + ". " + ss[1] + ". ";

            if (!string.IsNullOrEmpty(blockName) && !string.IsNullOrEmpty(subsectionName))
                subjectIndexDecoding += blockName + subsectionName + ". ";
            if (s.Length > 2)
                if (s[2].ToLower() == "дв")
                    subjectIndexDecoding += "Дисциплины по выбору.";
            this.SubjectIndexDecoding = subjectIndexDecoding;
        }

        public void SetCompetencies(IXLWorksheet workSheet, int row, Dictionary<string, string> compDic)
        {
            var resultList = new List<string>();
            var competenciesList = workSheet.Cell(row, workSheet.ColumnsUsed().Count()).Value.ToString().Split(';', ' ').ToList();
            foreach (var item in competenciesList)
                if (!string.IsNullOrEmpty(item))
                    if (compDic.ContainsKey(item))
                        resultList.Add($"{item}" + " -" + compDic[item]);
            this.Competencies = "\t" + string.Join(";\n\t", resultList) + ".";
        }
    }
}
