using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.Json;
using System.Text.Encodings.Web;
using System.Text.Unicode;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.EntityFrameworkCore;

namespace Json.Da
{
    class Program
    {
        static List<Employee> GenerateList()
        {
            var empList = new List<Employee>();
            var xlApp = new Excel.Application();
            var xlBook = xlApp.Workbooks.Open("C:\\Users\\PC\\Desktop\\C#\\Проекты\\Json.Da\\Json.Da\\bin\\Debug\\Сведения о преподах.xlsx");
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

        static void Main(string[] args)
        {
            
        }
    }
}
