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

            using (ApplicationContext db = new ApplicationContext())
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
                var syllist = AddSyllabus.GenerateSyllabus(disclist);
                foreach (var s in syllist)
                {
                    db.Syllabuses.Add(s);
                    db.SaveChanges();
                }
                Console.WriteLine("Учебные планы успешно сохранены");
                var mapWorkLoad = AddWorkload.GenerateNeedInfForWorkLoad();
                //var a = mapWorkLoad["Гутнова Алина Казбековна"][0];
                //var b = mapWorkLoad["Гутнова Алина Казбековна"][1];
                //if (a.Equals(b))
                //    Console.WriteLine("YES");
                //else
                //    Console.WriteLine("NO");
                //foreach (var item in mapWorkLoad)
                //{
                //    //Console.WriteLine("{0} - {1} - {2} - {3}", item.Key, item.Value[a].Item1, item.Value[a].Item2, item.Value[a].Item3);
                //    Console.WriteLine(item.Key);
                //    foreach (var tup in item.Value)
                //    {
                //        Console.WriteLine("{0} - {1} - {2}", tup.Item1, tup.Item2, tup.Item3);
                //    }

                //}

            }
        }
    }
}

