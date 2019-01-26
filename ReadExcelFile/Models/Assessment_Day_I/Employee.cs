using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models
{
    public class Employees
    {
        public string NoOfSalariedEmployees { get; set; }
        public string NoOfHourlyEmployee { get; set; }
        public string NoOfTemprorayWorkers { get; set; }
        public string NoOfSkilledEmployees{ get; set; }
        public int ActualTotalNoOfEmplyees{ get; set; }

    }
}