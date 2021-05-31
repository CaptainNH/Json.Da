using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Data;
using Newtonsoft.Json;
using Json.Da;

namespace Json.Da
{
    class AddWordTable
    {
        public static List<EducTechn> AddEducTechns()
        {
            string path = Environment.CurrentDirectory;
            var patJson = path + @"\..\..\Documents\Jsons\1educationalTechnologies.json";
            var jsonText = File.ReadAllText(patJson);
            var discJson = JsonConvert.DeserializeObject<DataSet>(jsonText);
            //Console.WriteLine(discJson.Tables["EducTechn"].Rows.Count);
            var listEduc = new List<EducTechn>();
            DataTable datTable = discJson.Tables["EducTechn"];
            foreach (DataRow row in datTable.Rows)
            {
                string theme = row["Theme"].ToString();
                string at = row["ActivityType"].ToString();
                string nh = row["NumberOfHours"].ToString();
                int nh2 = Convert.ToInt32(nh);
                string af = row["ActiveForms"].ToString();
                string iF = row["InteractiveForms"].ToString();
                var techn = new EducTechn
                {
                    Theme = theme,
                    ActivityType = at,
                    NumberOfHours = nh2,
                    ActiveForms = af,
                    InteractiveForms = iF
                };
                listEduc.Add(techn);
            }
            return listEduc;
        }


        public static List<DisccMap> AddDiscApp()
        {
            List<DisccMap> listDisc = new List<DisccMap>();
            string path = Environment.CurrentDirectory;
            var patJson = path + @"\..\..\Documents\Jsons\2disciplineMap.json";
            var jsonText = File.ReadAllText(patJson);
            var discJson = JsonConvert.DeserializeObject<DataSet>(jsonText);
            DataTable datTable = discJson.Tables["DiscMap"];
            foreach (DataRow row in datTable.Rows)
            {
                var discMap = new DisccMap
                {
                    DiscQuestion = row["DiscQuestion"].ToString(),
                    Lection = Convert.ToInt32(row["Lection"].ToString()),
                    Practice = Convert.ToInt32(row["Practice"].ToString()),
                    Content = row["Content"].ToString(),
                    Hours = Convert.ToInt32(row["Hours"].ToString()),
                    FormsControl = row["FormsControl"].ToString(),
                    Min = Convert.ToInt32(row["Min"].ToString()),
                    Max = Convert.ToInt32(row["Max"].ToString()),
                    Literature = row["Literature"].ToString()
                };
                listDisc.Add(discMap);
            }

            return listDisc;
        }


        public static List<ResultMark> AddResultM()
        {
            var listResultMark = new List<ResultMark>();
            string path = Environment.CurrentDirectory;
            string pathJson = path + @"\..\..\Documents\Jsons\3theAresultingAssessment.json";
            var jsonText = File.ReadAllText(pathJson);
            var discJson = JsonConvert.DeserializeObject<DataSet>(jsonText);
            var discTab = discJson.Tables["ResultMark"];
            foreach (DataRow row in discTab.Rows)
            {
                var ResMark = new ResultMark
                {
                    CurrentControl = row["CurrentControl"].ToString(),
                    FormsControl = row["FormsControl"].ToString(),
                    First = row["First"].ToString(),
                    Second = row["Second"].ToString(),
                    Third = row["Third"].ToString(),
                    Fourth = row["Fourth"].ToString()
                };
                listResultMark.Add(ResMark);
            }
            return listResultMark;
        }
    }
}