using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Json.Da
{
    class AddCompetencie
    {
        public static List<Competencie> CompList(List<Competencie> listComp,IXLWorksheet plan,IXLWorksheet competencie)
        {
            string directory = Syllabus.SetOnlyDirection(plan,"B18");
            int year = Syllabus.OnlyYear(plan, "T29");
            for (int i = 1; i <200; i++)
            {
                
                if (!string.IsNullOrEmpty(competencie.Cell("B" + i).Value.ToString()))
                {
                    var  comp= new Competencie { };
                    comp.Directory = directory;
                    comp.Year = year;
                    comp.Encryption = competencie.Cell("B" + i).Value.ToString();
                    comp.Decription = competencie.Cell("D" + i).Value.ToString();
                    listComp.Add(comp);
                }
            }
            return listComp;
        }

            public static List<Competencie> GenerateCompetencies()
        {
            var listComp = new List<Competencie>();
            string path = Environment.CurrentDirectory + @"\..\..\Documents\Бакалавриат\ПМ";//Путь до Debug


            var AllFiles = Directory.EnumerateFiles(path, "*.xls", SearchOption.AllDirectories);

            foreach (var pathFile in AllFiles)
            {
                // Console.WriteLine(pathFile);
                
                var xlBook = new XLWorkbook(pathFile);
                var xlTitle = xlBook.Worksheet("Титул");
                var xlPlan = xlBook.Worksheet("Компетенции");
               CompList(listComp, xlTitle, xlPlan);
            }
            return listComp;
         
        }
    }
}
