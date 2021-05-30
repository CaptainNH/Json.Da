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
using Newtonsoft.Json;

namespace Json.Da
{
    class Program
    {
        static void Main(string[] args)
        {
            //string path = Environment.CurrentDirectory;
            //string pathJson = path + @"\..\..\Documents\Jsons\0main.json";
            //var text = File.ReadAllText(pathJson);
            //Discipline disc = JsonConvert.DeserializeObject<Discipline>(text);
            //Console.WriteLine(disc.DiscMap.Count);
            ////foreach (var item in disc.DiscMap)
            ////{
            ////   Console.WriteLine(item);

            //var edic = disc.DiscMap;
            //foreach (var item in edic)
            //{
            //    var str = item.Split('*');
            //    foreach (var i in str)
            //    {
            //        Console.WriteLine(i);
            //    }
            //    Console.WriteLine("*********************************");
            //}


            using (ApplicationContext db = new ApplicationContext())
            {
                //    //var emplist = AddEmployee.GenerateList();
                //    //foreach (var e in emplist)
                //    //{
                //    //    db.Employees.Add(e);
                //    //    db.SaveChanges();
                //    //}
                //    //Console.WriteLine("Сотрудники успешно сохранены");

                var disclist = AddDiscipline.GenerateDisciplineList();

                foreach (var d in disclist)
                {
                    db.Disciplines.Add(d);
                    db.SaveChanges();
                }
                //Console.WriteLine("Предметы успешно сохранены");
                //    //var syllist = AddSyllabus.GenerateSyllabus(disclist);
                //    //foreach (var s in syllist)
                //    //{
                //    //    db.Syllabuses.Add(s);
                //    //    db.SaveChanges();
                //    //}
                //    //Console.WriteLine("Учебные планы успешно сохранены");
                //    //var wllist = AddWorkload.GenerateNeedInfForWorkLoad(emplist, disclist, syllist);
                //    //foreach (var wl in wllist)
                //    //{
                //    //    db.Workload.Add(wl);
                //    //    db.SaveChanges();
                //    //}
                //    //Console.WriteLine("Нагрузки успешно сохранены");
                //    //var compList = AddCompetencie.GenerateCompetencies();
                //    //foreach (var comp in compList)
                //    //{
                //    //    db.Competencie.Add(comp);
                //    //    db.SaveChanges();
                //    //}
            }
        }
    }
}

