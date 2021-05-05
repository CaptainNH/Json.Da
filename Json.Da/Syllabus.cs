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

        public int Standart { get; set; }

        public int Hours { get; set; }

        public Syllabus()
        {
            Direction = "none";
            Profile = "none";

        }

        public void DirectionAndProfile(IXLWorksheet workSheet)
        {
            string[] separators = new string[] {"Направленность программы", "Направление подготовки", "Профиль",
                "Профиль:", "Профили", "Направление", "Программа"}; //разделители направления и профиля
            var directionAndProfile = workSheet.Cell("B18").Value.ToString().Split(separators, StringSplitOptions.RemoveEmptyEntries); //Сплит по разделителям
            Direction = directionAndProfile[0].Trim(' ', ',', ':'); //Получить направление
            Profile = "";
            if (directionAndProfile.Length > 1)
                Profile = "Профиль: " + directionAndProfile[1].Trim(' ', ':'); //Получить профиль, если он есть            
        }
    }
}
