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
            var xlApp = new Microsoft.Office.Interop.Excel.Application();
            var xlBook = xlApp.Workbooks.Open(path + @"\..\..\Documents\Svedenia.xlsx");
            var xlSheet = xlBook.Worksheets["Сведения о преподавателях"];
            for (int i = 3; i <= 120; i++)
            {
                string[] fio = xlSheet.Cells[1][i].Value.Split();
                string surname = fio[0].Trim();
                string name = fio[1].Trim();
                string fathername = fio[2].Trim();
                string[] s = xlSheet.Cells[3][i].Value.Split(',');
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
            if (xlApp != null)
            {
                xlApp.Quit();
                Marshal.ReleaseComObject(xlSheet);
                Marshal.ReleaseComObject(xlApp);
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
                //db.Employees.AddRange(emplist);
                //db.Add(new Employee() { Name = "Стасян", Surname = "А.", Fathername = "Б.", Position = "Да", Rank = "Нет" });
                //db.SaveChanges();
                Console.WriteLine("Объекты успешно сохранены");
                // получаем объекты из бд и выводим на консоль
                //var emp = db.Employees.ToList();
                //Console.WriteLine("Список объектов:");
                //foreach (Employee e in emp)
                //{
                //    Console.WriteLine($"{e.Id}.{e.Surname} - {e.Name} - {e.Fathername} - {e.Position} - {e.Rank}");
                //}
            }
            //Console.Read();
        }
    }
}
