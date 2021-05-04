using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Json.Da
{
    class AddEmployee
    {
        static List<Employee> GenerateList()
        {
            string path = Environment.CurrentDirectory;
            var empList = new List<Employee>();
            var xlSheetSvedenia = OpenExcelFile(path + @"\..\..\Documents\Svedenia.xlsx", "Сведения о преподавателях");
            var xlSheetNagruzki = OpenExcelFile(path + @"\..\..\Documents\Nagruzki.xlsx", "Сводное поручение");
            var range1 = xlSheetSvedenia.Range("A3:C120");
            var range2 = xlSheetNagruzki.Range("B13:B19");
            foreach (var row in range1.Rows())
            {
                string[] fio = row.Cell(1).GetValue<string>().Trim().Split();
                string surname = fio[0].Trim();
                string name = fio[1].Trim();
                string fathername = fio[2].Trim();
                string fioSearch = string.Join(" ", surname + name + fathername);
                if (range2.Contains(fioSearch))
                {
                    
                }
                string[] s = row.Cell(3).GetValue<string>().Split(',');
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
                    Rank = rank
                };
                empList.Add(prepod);
            }            
            return empList;
        }

        public static void AddToDB()
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                var emplist = GenerateList();
                foreach (Employee e in emplist)
                {
                    db.Employees.Add(e);
                    db.SaveChanges();
                }
                Console.WriteLine("Объекты успешно сохранены");
            }
        }
    }
}
