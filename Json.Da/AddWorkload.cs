using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace Json.Da
{
    class AddWorkload
    {
        private static Dictionary<string, List<Tuple<string, string, string>>> GenerateWorkloadDic(IXLWorksheets xlLists)
        {
            var mapWorkLoad = new Dictionary<string, List<Tuple<string, string, string>>>();
            foreach (var worksheet in xlLists)
            {
                string emplName = worksheet.Cell("A4").Value.ToString();//Имя преподавателя
                if (!mapWorkLoad.ContainsKey(emplName) && !String.IsNullOrEmpty(emplName))
                {
                    var listTup = new List<Tuple<string, string, string>>();
                    for (int i = 13; i < 150; i++)
                    {
                        var cell = worksheet.Cell("B" + i);
                        if (!cell.Style.Font.Bold && !String.IsNullOrEmpty(cell.Value.ToString()))
                        {
                            var syllabus = worksheet.Cell("B" + i).Value.ToString();//Учебный план
                            var disc = worksheet.Cell("E" + i).Value.ToString();//Дисциплина
                            var sem = worksheet.Cell("F" + i).Value.ToString();//Семестр 
                            var tup = Tuple.Create(syllabus, disc, sem);
                            listTup.Add(tup);
                        }
                    }
                    mapWorkLoad[emplName] = listTup;
                }
            }
            return mapWorkLoad;
        }

        public static List<Workload> GenerateNeedInfForWorkLoad(List<Employee> emplist, List<Discipline> disclist, List<Syllabus> syllist)
        { 
            string path = Environment.CurrentDirectory + @"\..\..\Documents\Nagruzki.xlsx";
            XLWorkbook xlBook = new XLWorkbook(path);
            var xlLists = xlBook.Worksheets;
            var list = xlBook.Worksheet("Сводное поручение");
            var year = list.Cell("A9").GetValue<string>().Split()[3];
            var mapWorkLoad = GenerateWorkloadDic(xlLists);
            List<Workload> workloadList = new List<Workload>();
            foreach (var item in mapWorkLoad)
            {
                Employee emp = emplist
                    .Where(x => string.Join(" ", new string[] { x.Surname, x.Name, x.Fathername }) == item.Key)
                    .FirstOrDefault();
                foreach (var el in item.Value)
                {
                    Discipline disc = disclist
                        .Where(x => x.Name == el.Item2)
                        .FirstOrDefault();
                    List<Syllabus> syl = syllist
                        .Where(x => x.Direction.Split()[0] == el.Item1.Split(' ', '-', '_')[1])
                        .Where(x => disc != null && x.Predmet.Name == disc.Name)
                        .ToList();
                    foreach (var s in syl)
                    {
                        var wl = new Workload()
                        {
                            Year = year,
                            Employee = emp,
                            Plan = s
                        };
                        workloadList.Add(wl);
                    }
                }
            }
            return workloadList;
        }
  
    }

}
