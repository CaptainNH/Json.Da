using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Json.Da
{
    class AddSyllabus
    {
       


        public static List<Syllabus> F()
        {
            var listSyllabus = new List<Syllabus>();

            string path = Environment.CurrentDirectory;//Путь до Debug
            string pathPm1 = path + @"\..\..\Documents\Бакалавриат\ПМ\B010302-20-1-ПМ.xlsx";//Путь до ПМ-2020
            var xlBookPm1 = new XLWorkbook(pathPm1);
            var xlPM1Plan = xlBookPm1.Worksheet("План");
            var xlPM1Title = xlBookPm1.Worksheet("Титул");
            //Console.WriteLine(xlPM1Plan.LastColumn());
            var firstColumn = 'Q'-'A'+1;
            var lastColumn = ('D'-'A'+1)*('Z'-'A'+1);
            Console.WriteLine(xlPM1Plan.Cell(2,lastColumn));
            for (int i = firstColumn; i < lastColumn; i+=7)
            {
                var syllabus = new Syllabus();
                
                syllabus.Year = Convert.ToInt32(xlPM1Title.Cell("T29").Value.ToString());
                syllabus.DirectionAndProfile(xlPM1Title);



                syllabus.Semester = Convert.ToInt32(xlPM1Plan.Cell(2, i).Value.ToString().Split()[1]);
            }

            return listSyllabus;
                
        }

    }


}








//var lastColumn = xlPM1Plan.Cell("CZ2");
//var firstColumn = xlPM1Plan.Cell("Q2");
//var discRange = xlPM1Plan.Range(firstColumn, lastColumn);

//foreach (var item in discRange.Cells())
//{
//    if (!string.IsNullOrEmpty(item.Value.ToString()))
//    {
//        Console.WriteLine(item);
//    }
//}