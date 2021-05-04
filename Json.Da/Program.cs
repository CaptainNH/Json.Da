using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.Json;
using System.Text.Encodings.Web;
using System.Text.Unicode;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.EntityFrameworkCore;

namespace Json.Da
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello");
            var hashSet = AddDiscipline.GenerateHash();
            int a = 0;
            foreach (var item in hashSet)
            { 
                Console.WriteLine(a++ +" "+ item.Name+" "+item.Competencies);
            }
            Console.WriteLine("Bye Bye");
            AddEmployee.AddToDB();

        }
    }
}

