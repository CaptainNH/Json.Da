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
using System.Data;

namespace Json.Da
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = Environment.CurrentDirectory;
            string pathJson = path + @"\..\..\Documents\Jsons\3theAresultingAssessment.json";
            var text = File.ReadAllText(pathJson);
            var disc = JsonConvert.DeserializeObject<DataSet>(text);
            var discTab = disc.Tables["ResultMark"];
            foreach (DataRow row in discTab.Rows)
            {
                Console.WriteLine(row["Second"]);
            }



            using (ApplicationContext db = new ApplicationContext())
            {
                var emplist = AddEmployee.GenerateList();
                foreach (var e in emplist)
                {
                    db.Employees.Add(e);
                    db.SaveChanges();
                }
                Console.WriteLine("Сотрудники успешно сохранены");
                var educList = AddWordTable.AddEducTechns();
                foreach (var ed in educList)
                {
                    db.EducTechns.Add(ed);
                    db.SaveChanges();
                }
                Console.WriteLine("Таблица успешно сохранена");
                var resMarkList = AddWordTable.AddResultM();
                foreach (var rm in resMarkList)
                {
                    db.ResultMarks.Add(rm);
                    db.SaveChanges();
                }
                Console.WriteLine("Таблица аубуба");
                var discMapList = AddWordTable.AddDiscApp();
                foreach (var dm in discMapList)
                {
                    db.DisccMaps.Add(dm);
                    db.SaveChanges();
                }
                Console.WriteLine("Таблица аубуба");
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
                var wllist = AddWorkload.GenerateNeedInfForWorkLoad(emplist, disclist, syllist);
                foreach (var wl in wllist)
                {
                    db.Workload.Add(wl);
                    db.SaveChanges();
                }
                Console.WriteLine("Нагрузки успешно сохранены");
            }
        }
    }
}

