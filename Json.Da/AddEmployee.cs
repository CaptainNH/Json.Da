using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace Json.Da
{
    class AddEmployee
    {
        static IXLWorksheet Svedenia;
        static IXLWorksheet Nagruzki;

        static IXLWorksheet OpenExcelFile(string filePath, string sheetName)
        {
            var xlBook = new XLWorkbook(filePath);
            var xlSheet = xlBook.Worksheet(sheetName);
            return xlSheet;
        }

        static double FindRate(string fio, IXLRange range, ref string chair)
        {
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            double rate = 0;
            foreach (var row in range.Rows())
            {
                string s = row.FirstCell().GetString().Trim();
                if (s == fio)
                {
                    rate = Convert.ToDouble(row.LastCell().GetString().Split(',')[1].Trim().Split()[0]);
                    chair = Nagruzki.Cell("C2").GetString();
                }
            }
            return rate;
        }

        public static List<Employee> GenerateList()
        {
            string path = Environment.CurrentDirectory;
            var empList = new List<Employee>();
            Svedenia = OpenExcelFile(path + @"\..\..\Documents\Svedenia.xlsx", "Сведения о преподавателях");
            Nagruzki = OpenExcelFile(path + @"\..\..\Documents\Nagruzki.xlsx", "Сводное поручение");
            var range1 = Svedenia.Range("A3:C120");
            var range2 = Nagruzki.Range("B13:I19");
            foreach (var row in range1.Rows())
            {
                string[] fio = row.FirstCell().GetString().Trim().Split();
                string surname = fio[0].Trim();
                string name = fio[1].Trim();
                string fathername = fio[2].Trim();
                string fioSearch = $"{surname} {name} {fathername}";
                string chair = "-";
                double rate = FindRate(fioSearch, range2, ref chair);
                string[] s = row.Cell(3).GetString().Split(',');
                string pos = s[0].Trim().Split(new string[] { "Должность" }, StringSplitOptions.None)[1].Trim(new char[] { ' ', '-', '–' });
                string rank = "-";
                if (s.Length == 3)
                {
                    if (!s[2].ToLower().Contains("отсутствует"))
                        rank = s[2].Trim();
                }
                else
                    if (!s[3].ToLower().Contains("отсутствует"))
                    rank = s[3].Trim();
                var prepod = new Employee()
                {
                    Surname = surname,
                    Name = name,
                    Fathername = fathername,
                    Position = pos,
                    Rank = rank,
                    Rate = rate,
                    Chair = chair                    
                };
                empList.Add(prepod);
            }            
            return empList;
        }
    }
}
