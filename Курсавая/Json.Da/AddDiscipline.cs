using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Data;
using Newtonsoft.Json;
using System.Data;

namespace Json.Da
{
    class Json 
    {
        
    }
    class AddDiscipline
    {
        public static Dictionary<string, string> discMap = new Dictionary<string, string>();
        static void AddToHash(IXLWorksheet workSheet)
        {
            var discRange = workSheet.Range("C6", "C130");
            foreach (var item in discRange.Cells())
            {
                if (!string.IsNullOrEmpty(item.Value.ToString()) && !item.Style.Font.Bold)
                {
                    int rowNumb = item.Address.RowNumber;
                    string key = item.Value.ToString();
                    if (!discMap.ContainsKey(key))
                    {
                        discMap[key] = workSheet.Cell("BW" + rowNumb.ToString()).Value.ToString(); ;
                    }
                }
            }

        }

        public static List<Discipline> GenerateDisciplineList()
        {

            string path = Environment.CurrentDirectory;//Путь до Debug
            string pathPm1 = path + @"\..\..\Documents\Бакалавриат\ПМ\B010302-20-1-ПМ.xlsx";//Путь до ПМ-2020
            string pathPm2 = path + @"\..\..\Documents\Бакалавриат\ПМ\B010302-20-2-ПМ  МатМод Дзанагова.plx.xlsx";//Путь до ПМ-2021
            string pathPm3MathEconom = path + @"\..\..\Documents\Бакалавриат\ПМ\B010302-20-3-ПМ _МатЭкон Дзанагова.plx.xlsx";
            string pathPm3MathMod = path + @"\..\..\Documents\Бакалавриат\ПМ\B010302-20-3-ПМ_МатМод Дзанагова.plx.xlsx";
            string pathPm4 = path + @"\..\..\Documents\Бакалавриат\ПМ\B010302-20-4-ПМ+.plx.xlsx";
            var xlBookPm1 = new XLWorkbook(pathPm1);
            var xlBookPm2 = new XLWorkbook(pathPm2);
            var xlBookPm3MathEconom = new XLWorkbook(pathPm3MathEconom);
            var xlBookPm3MathMod = new XLWorkbook(pathPm3MathMod);
            var xlBookPm4 = new XLWorkbook(pathPm4);
            var xlPM1Plan = xlBookPm1.Worksheet("План");
            var xlPM2Plan = xlBookPm2.Worksheet("План");
            var xlPM3MathEconomPlan = xlBookPm3MathEconom.Worksheet("План");
            var xlPM3MathModPlan = xlBookPm3MathMod.Worksheet("План");
            var xlPM4Plan = xlBookPm4.Worksheet("План");
            AddToHash(xlPM1Plan);
            AddToHash(xlPM2Plan);
            AddToHash(xlPM3MathEconomPlan);
            AddToHash(xlPM3MathModPlan);
            AddToHash(xlPM4Plan);
            var patJson = path + @"\..\..\Documents\Jsons\0main.json";
            var jsonText= File.ReadAllText(patJson);
            var discList = new List<Discipline>();
            var discJson = JsonConvert.DeserializeObject<Discipline>(jsonText);
            
            foreach (var item in discMap)
            {
                var discipline = new Discipline
                {
                    Name = item.Key,
                    Competencies = item.Value,
                    NamePat = discJson.NamePat,
                    Koi = discJson.Koi,
                    Date= discJson.Date,
                    DisciplineTarget= discJson.DisciplineTarget,
                    OPOP= discJson.OPOP,
                    Know= discJson.Know,
                    BeAbleTo= discJson.BeAbleTo,
                    Own= discJson.Own,
                    ControlTasks= discJson.ControlTasks,
                    TestTasks= discJson.TestTasks,
                    QuestionForTest= discJson.QuestionForTest,
                    InformationSupportOfDiscipline= discJson.InformationSupportOfDiscipline,
                    LogisticsOfTheDiscipline= discJson.LogisticsOfTheDiscipline,
                    UpdateSheet= discJson.UpdateSheet,
                    EducTechn=discJson.EducTechn,
                    DiscMap=discJson.DiscMap
                };
                discList.Add(discipline);
            }
            return discList;
        }
    }
}
