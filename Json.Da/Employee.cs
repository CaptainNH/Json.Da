using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Json.Da
{
    class Employee
    {
        public int Id { get; set; }

        public string Surname { get; set; }//Фамилия

        public string Name { get; set; }//Имя

        public string Fathername { get; set; }//Отчество

        public string Position { get; set; }//Должность

        public string Rank { get; set; }//Звание

        public double Rate { get; set; }//Показатель

        public string Chair { get; set; }//Кафедра
    }
}
