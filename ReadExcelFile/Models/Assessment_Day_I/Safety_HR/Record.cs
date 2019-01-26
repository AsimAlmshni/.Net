using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Safety_HR
{
    public class Record
    {
        public string PriorYearTotal { get; set; }
        public string PYMA { get; set; }
        public string CYMA { get; set; }
        public string YTDForecast { get; set; }
    }
}