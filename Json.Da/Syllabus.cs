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
        //Поле  для базы данных
        public int Id { get; set; }//

        //Вспомогательное поле
        private bool IsGraduateSchool { get; set; } //Показывает это аспирантуриа или нет

        //Основные поля
        public int Year { get; set; } //Год начала обучения
        public string Direction { get; set; } //Направле обучения
        public string Profile { get; set; } //Профиль
        public string StudyProgram { get; set; } //Программа обучения
        public string Standart {get; set; } //Стандарт
        public string Protocol { get; set; } //Протокоп
        public string EdForm { get; set; } //Форма обучения
        public string DirectionAbbreviation { get; set; } //Аббревиатура направления обучения
        public string Director { get; set; } //Директор
        public string Position { get; set; } //Должность
        public string SubjectName { get; set; } //Название дисциплины
        public Discipline Predmet { get; set; } //Содержит неизменяемую информацию о дисциплине
        public string Semester { get; set; } //Номер семестра
        public string Course { get; set; } //Номер курса
        public int CreditUnits { get; set; } //Зачётные единицы
        public string Hours { get; set; } //Общее количество часов
        public string CourseWork { get; set; } //Наличие курсовой работы
        public string SumIndependentWork { get; set; } //Сумма часов самостоятельной работы
        public string InteractiveWatch { get; set; } //Сумма интерактивных часов
        public bool Test { get; set; } //Наличие зачёта
        public bool Exam { get; set; } //Наличие экзамена
        public bool Consulting { get; set; } //Наличие консультаций
        public int Lectures { get; set; } //Количество часов лекций за семестр
        public int LaboratoryExercises { get; set; } //Количество часов лабораторных работ за семестр
        public int Workshops { get; set; } //Количество часов практик за семестр
        public double IndependentWorkBySemester { get; set; } //Количество часов самостоятельных работ за семестр
        public string TypesOfLessons { get; set; } //Типы занятий за семестр
        public double AuditoryLessons { get; set; } //Общее количество аудиторных занятий
        public string SubjectIndex { get; set; } //Индекс дисциплины
        public string SubjectIndexDecoding { get; set; } //Расшифровка индекса дисциплины
        public string Competencies { get; set; } //Компетенции

        public Syllabus()
        {
            Semester = "-";
        }

        public Syllabus
            (List<Discipline> predmetlist, 
            IXLWorksheet workSheetTitle, IXLWorksheet workSheetPlan, 
            Dictionary<string, string> compDic, int row, int column)
        {
            //Запонение из титульного листа
            SetYear(workSheetTitle, "T29"); //Извлкает год начала обучения
            SetDirectionAndProfile(workSheetTitle, "B18"); //Извлкает направление обучения и профиль
            SetStudyProgram(workSheetTitle, "F14"); //Извлкает программу обучения
            SetStandart(workSheetTitle, "T31"); //Извлкает стандарт
            SetProtocol(workSheetTitle, "A13"); //Извлкает протокол
            SetEdForm(workSheetTitle, "A31", "A30"); //Извлкает форму обучения
            SetDirectionAbbreviation(workSheetTitle, "B18"); //Извлкает аббревиатуру направления обучения
            SetDirestor("А.М. Дигурова", "Б.В. Туаева", "Л.А. Агузарова"); //Устанавливает директора
            SetPosition("Проректор по УР", "Проректор по научной деятельности", "Первый проректор"); //Устанавливает должность
            //Запонение из листа план
            SetSemester(workSheetPlan, column); //Извлкает номер семестра
            SetAuditoryLessons(workSheetPlan, row); //Извлкает сумму аудиторных занятий
            SetCourseWork(workSheetPlan, row); //Устанавливает наличие курсовой работы
            SetCreditUnits(workSheetPlan, row); //Извлкает зачётные еденицы
            SetExam(workSheetPlan, row); //Устанавливает наличие экзаменя
            SetHours(workSheetPlan, row); //Извлкает общее количество часов обучения
            SetSubjectName(workSheetPlan, row); //Извлкает название предмета
            SetInteractiveWatch(workSheetPlan, row); //Извлкает количество интеративных часов
            SetLaboratoryExercises(workSheetPlan, row, column); //Извлкает количество часов лабораторных работ
            SetLestures(workSheetPlan, row, column); //Извлкает количество часов лекций
            SetSumIndependentWork(workSheetPlan, row); //Извлкает общее количество часов самостоятельной работы
            SetTests(workSheetPlan, row); //Устанавливает наличие зачёта
            SetWorkshops(workSheetPlan, row, column); //Извлкает количество часов практик
            SetCourse(); //Устанавливает номер курса
            SetTypesOfLessons(); //Устанавливает типы аудиторных занятий
            SetCompetencies(workSheetPlan, row, compDic); //Устанавливает компетенции
            SetSubjectIndex(workSheetPlan, row); //Извлкает индекс дисциплины
            DecodeSubjectIndex(workSheetPlan, row); //Расшифровывает индекс дисциплины
            SetIndependentWorkBySemester(workSheetPlan, row, column); //Извлкает количество часо самостоятельной работы за семестр
            SetConsulting(workSheetPlan, row); //Устанавливает наличие консультаций
            Predmet = predmetlist.Find(item => item.Name == this.SubjectName); //Сопоставляет каждой дисциплине класс Discipline
        }

        public void SetYear(IXLWorksheet workSheet, string cellName)
        {
            if (!string.IsNullOrEmpty(workSheet.Cell(cellName).Value.ToString()))// Проверка на пустоту
                this.Year = Convert.ToInt32(workSheet.Cell(cellName).Value.ToString());// присваивание значения
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
                i--;//Двигается вверх пока не найдёт пустую клетку в столбце "Наименование"
            subsectionName = workSheet.Cell(i, 1).Value.ToString().Trim(' ');//Извлекает название части

            while (workSheet.Cell(i, 1).Value.ToString().Split()[0].ToLower() != "блок")
                i--;//Двигается вверх пока не найдёт клетку, в которой написано слово блок" в столбце "Считать в плане"
            string[] ss = workSheet.Cell(i, 1).Value.ToString().Trim(' ').Split('.');
            blockName = ss[0] + ". " + ss[1] + ". ";//Извлекает название и номер блока

            if (!string.IsNullOrEmpty(blockName) && !string.IsNullOrEmpty(subsectionName))
                subjectIndexDecoding += blockName + subsectionName + ". ";//Объеденяет название блока и названи части
            if (s.Length > 2)
                if (s[2].ToLower() == "дв")
                    subjectIndexDecoding += "Дисциплины по выбору.";//Добавляет "дисциплины по выбору" если они есть
            this.SubjectIndexDecoding = subjectIndexDecoding;
        }

        public void SetCompetencies(IXLWorksheet workSheet, int row, Dictionary<string, string> compDic)
        {
            var resultList = new List<string>();
            var competenciesList = workSheet.Cell(row, 
                workSheet.ColumnsUsed().Count()).Value.ToString().Split(';', ' ').ToList();
            //Извлекает компетенции дисциплины
            foreach (var item in competenciesList)
                if (!string.IsNullOrEmpty(item))
                    if (compDic.ContainsKey(item))
                        resultList.Add($"{item}" + " -" + compDic[item]);//Находит такие же в словаре
            this.Competencies = "\t" + string.Join(";\n\t", resultList) + ".";//Переводит в строку
        }
    }
}
