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
                            syllabus.SetStudyProgram(workSheetTitle, "F14");
                            syllabus.SetStandart(workSheetTitle, "T31");
                            syllabus.SetProtocol(workSheetTitle, "A13");
                            syllabus.SetEdForm(workSheetTitle, "A31", "A30");
                            syllabus.SetDirectionAbbreviation(workSheetTitle, "B18");
                            syllabus.SetDirestor("А.М. Дигурова", "Б.В. Туаева", "Л.А. Агузарова");
                            syllabus.SetDirestor("Проректор по УР", "Проректор по научной деятельности", "Первый проректор");

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
                            syllabus.SetCourse();
                            syllabus.SetTypesOfLessons();

                            listSyllabus.Add(syllabus);
                        }
                    }
            }
        }

        public static List<Syllabus> GenerateSyllabus(List<Discipline> predmetlist)
        {
            var listSyllabus = new List<Syllabus>();
            string path = Environment.CurrentDirectory + @"\..\..\Documents\Бакалавриат\ПМ";//Путь до Debug


            var AllFiles = Directory.EnumerateFiles(path, "*.xls", SearchOption.AllDirectories);
            foreach (var pathFile in AllFiles)
            {
                Console.WriteLine(pathFile);
                var xlBook = new XLWorkbook(pathFile);
                var xlTitle = xlBook.Worksheet("Титул");
                var xlPlan = xlBook.Worksheet("План");
                FileProcessing(predmetlist, listSyllabus, xlTitle, xlPlan);
            }
            return listSyllabus;               
        }
    }
}