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
        public static void FileProcessing(List<Discipline> predmetlist, List<Syllabus> listSyllabus, IXLWorksheet workSheetTitle, IXLWorksheet workSheetPlan)
        {
            var firstColumn = 'Q' - 'A' + 1;
            var lastColumn = ('D' - 'A' + 1) * ('Z' - 'A' + 1);
            for (int r = 6; r < 150; r++)
            {
                var subjectName = workSheetPlan.Cell(r, 3);
                if (!string.IsNullOrEmpty(subjectName.Value.ToString()) && !subjectName.Style.Font.Bold)
                    for (int c = firstColumn; c < lastColumn; c += 7)
                    {
                        if (!string.IsNullOrEmpty(workSheetPlan.Cell(2, c).Value.ToString()))
                        {
                            var syllabus = new Syllabus();

                            syllabus.Predmet = predmetlist.Find(item => item.Name == subjectName.Value.ToString());

                            syllabus.SetYear(workSheetTitle, "T29");
                            syllabus.SetDirectionAndProfile(workSheetTitle, "B18");

                            syllabus.SetSemester(workSheetPlan, c);
                            syllabus.SetAuditoryLessons(workSheetPlan, r);
                            syllabus.SetCourseWork(workSheetPlan, r);
                            syllabus.SetCreditUnits(workSheetPlan, r);
                            syllabus.SetExam(workSheetPlan, r);
                            syllabus.SetHours(workSheetPlan, r);
                            syllabus.SetSubjectName(workSheetPlan, r);
                            syllabus.SetInteractiveWatch(workSheetPlan, r);
                            syllabus.SetLaboratoryExercises(workSheetPlan, r, c);
                            syllabus.SetLestures(workSheetPlan, r, c);                            
                            syllabus.SetSumIndependentWork(workSheetPlan, r);
                            syllabus.SetTests(workSheetPlan, r);
                            syllabus.SetWorkshops(workSheetPlan, r, c);
                            syllabus.SetTypesOfLessons();

                            listSyllabus.Add(syllabus);
                        }
                    }
            }
        }

        public static List<Syllabus> GenerateSyllabus(List<Discipline> predmetlist)
        {
            var listSyllabus = new List<Syllabus>();
            string path = Environment.CurrentDirectory;//Путь до Debug

            string pathPm1 = path + @"\..\..\Documents\Бакалавриат\ПМ\B010302-20-1-ПМ.xlsx";//Путь до ПМ-2020
            var xlBookPm1 = new XLWorkbook(pathPm1);
            var xlPM1Title = xlBookPm1.Worksheet("Титул");
            var xlPM1Plan = xlBookPm1.Worksheet("План");
            FileProcessing(predmetlist, listSyllabus, xlPM1Title, xlPM1Plan);
           
            string pathPm2 = path + @"\..\..\Documents\Бакалавриат\ПМ\B010302-20-2-ПМ  МатМод Дзанагова.plx.xlsx";//Путь до ПМ-2021
            var xlBookPm2 = new XLWorkbook(pathPm2);
           var xlPM2Title = xlBookPm2.Worksheet("Титул");
            var xlPM2Plan = xlBookPm2.Worksheet("План");
            FileProcessing(predmetlist, listSyllabus, xlPM2Title, xlPM2Plan);

            string pathPm3MathEconom = path + @"\..\..\Documents\Бакалавриат\ПМ\B010302-20-3-ПМ _МатЭкон Дзанагова.plx.xlsx";
            var xlBookPm3MathEconom = new XLWorkbook(pathPm3MathEconom);
            var xlPM3MathEconomTitle = xlBookPm3MathEconom.Worksheet("Титул");
            var xlPM3MathEconomPlan = xlBookPm3MathEconom.Worksheet("План");
            FileProcessing(predmetlist, listSyllabus, xlPM3MathEconomTitle, xlPM3MathEconomPlan);

           string pathPm3MathMod = path + @"\..\..\Documents\Бакалавриат\ПМ\B010302-20-3-ПМ_МатМод Дзанагова.plx.xlsx";
            var xlBookPm3MathMod = new XLWorkbook(pathPm3MathMod);
            var xlPM3MathModTitlee = xlBookPm3MathMod.Worksheet("Титул");
            var xlPM3MathModPlan = xlBookPm3MathMod.Worksheet("План");
            FileProcessing(predmetlist, listSyllabus, xlPM3MathModTitlee, xlPM3MathModPlan);

            string pathPm4 = path + @"\..\..\Documents\Бакалавриат\ПМ\B010302-20-4-ПМ+.plx.xlsx";
            var xlBookPm4 = new XLWorkbook(pathPm4);
            var xlPM4Title = xlBookPm4.Worksheet("Титул");
            var xlPM4Plan = xlBookPm4.Worksheet("План");
            FileProcessing(predmetlist, listSyllabus, xlPM4Title, xlPM4Plan);

            return listSyllabus;               
        }
    }
}