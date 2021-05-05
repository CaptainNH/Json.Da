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
        static void Main(string[] args)
        {

            AddSyllabus.F();
            /*using (ApplicationContext db = new ApplicationContext())
            {
                var emplist = AddEmployee.GenerateList();
                foreach (var e in emplist)
                {
                    db.Employees.Add(e);
                    db.SaveChanges();
                }
                Console.WriteLine("Сотрудники успешно сохранены");
                var disclist = AddDiscipline.GenerateDisciplineList();
                foreach (var d in disclist)
                {
                    db.Disciplines.Add(d);
                    db.SaveChanges();
                }
                Console.WriteLine("Предметы успешно сохранены");
            }*/
        }
    }
}

