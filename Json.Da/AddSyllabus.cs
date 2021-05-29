using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Json.Da
{
    class AddSyllabus
    {
        public static void FileProcessing(List<Discipline> predmetlist, List<Syllabus> listSyllabus, IXLWorksheet workSheetTitle, IXLWorksheet workSheetPlan, IXLWorksheet workSheetComp)
        {
            var compDic = CreateCompDic(workSheetComp);
            var firstColumn = 'Q' - 'A' + 1;
            var lastColumn = ('D' - 'A' + 1) * ('Z' - 'A' + 1);
            var firstRow = 6;
            var lastRow = workSheetPlan.RowsUsed().Count();
            for (int r = firstRow; r < lastRow; r++)
            {
                var subjectName = workSheetPlan.Cell(r, 3);
                if (!string.IsNullOrEmpty(subjectName.Value.ToString()) && !subjectName.Style.Font.Bold)
                    for (int c = firstColumn; c < lastColumn; c += 7)
                        if (!string.IsNullOrEmpty(workSheetPlan.Cell(2, c).Value.ToString()) 
                            && !string.IsNullOrEmpty(workSheetPlan.Cell(r, c+1).Value.ToString()))
                            listSyllabus.Add(
                                new Syllabus(predmetlist, listSyllabus, workSheetTitle, workSheetPlan, workSheetComp, compDic, r, c)
                                );
            }
        }

        public static List<Syllabus> GenerateSyllabus(List<Discipline> predmetlist)
        {
            var listSyllabus = new List<Syllabus>();
            string path = Environment.CurrentDirectory + @"\..\..\Documents\Бакалавриат\ПМ";//Путь до Debug
            //string path = Environment.CurrentDirectory + @"\..\..\Documents\Аспирантура";//Путь до Debug

            var AllFiles = Directory.EnumerateFiles(path, "*.xls", SearchOption.AllDirectories);
            foreach (var pathFile in AllFiles)
            {
                Console.WriteLine(pathFile);
                var xlBook = new XLWorkbook(pathFile);
                var xlTitle = xlBook.Worksheet("Титул");
                var xlPlan = xlBook.Worksheet("План");
                var xlComp = xlBook.Worksheet("Компетенции");
                FileProcessing(predmetlist, listSyllabus, xlTitle, xlPlan, xlComp);
            }
            return listSyllabus;               
        }

        public static Dictionary<string, string> CreateCompDic(IXLWorksheet competencie)
        {
            Dictionary<string, string> compDic = new Dictionary<string, string>();
            for (int i = 1; i < 200; i++)
                if (!string.IsNullOrEmpty(competencie.Cell("B" + i).Value.ToString()))
                    compDic.Add(competencie.Cell("B" + i).Value.ToString(),
                         competencie.Cell("D" + i).Value.ToString());
            return compDic;
        }

    }
}