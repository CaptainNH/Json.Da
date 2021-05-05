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
            Console.WriteLine(xlPM1Plan.LastRowUsed());
            Console.WriteLine(xlPM1Plan.LastRow().RangeAddress);
            for (int r = 6; r < 150; r++)
            {
                var subjectName = xlPM1Plan.Cell(r, 3);
                if (!string.IsNullOrEmpty(subjectName.Value.ToString()) && !subjectName.Style.Font.Bold)
                    for (int c = firstColumn; c < lastColumn; c += 7)
                    {
                        if (!string.IsNullOrEmpty(xlPM1Plan.Cell(2, c).Value.ToString()))
                        {
                            var syllabus = new Syllabus();

                            syllabus.Year = Convert.ToInt32(xlPM1Title.Cell("T29").Value.ToString());
                            syllabus.SetDirectionAndProfile(xlPM1Title);

                            syllabus.SubjectName = subjectName.Value.ToString();
                            //syllabus.Predmet = AddDiscipline.discMap[syllabus.SubjectName];
                            syllabus.Semester = Convert.ToInt32(xlPM1Plan.Cell(2, c).Value.ToString().Split()[1]);
                        }
                    }
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