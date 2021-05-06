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
        //  public static List<Workload> GenerateWorkLoad()/*List<Employee>,List<Syllabus>*/
        public static Dictionary<string, List<Tuple<string,string,string>>> GenerateNeedInfForWorkLoad()
        { 
            string path = Environment.CurrentDirectory + @"\..\..\Documents\Nagruzki.xlsx";

            XLWorkbook xlBook = new XLWorkbook(path);

            var xlLists = xlBook.Worksheets;

            var mapWorkLoad= new Dictionary<string, List<Tuple<string, string,string>>>();

            foreach (var worksheet in xlLists)
            {
                string emplName = worksheet.Cell("A4").Value.ToString();//Имя преподавателя
                if (!mapWorkLoad.ContainsKey(emplName) && !String.IsNullOrEmpty(emplName))
                {

                    // var discRange = worksheet.Range("A13", "A100");
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
                    //var list1 = new List<Tuple<string, string, string>>();
                    //for (int i = 1; i < listTup.Count; i++)
                    //{
                    //    if (!listTup[i-1].Equals(listTup[i]))
                    //    {
                    //        list1.Add(listTup[i]);
                    //    }
                    //} 
                    mapWorkLoad[emplName] = listTup;
                }
            }
            return mapWorkLoad;
        }
  
    }

}
