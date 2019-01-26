using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models
{
    public class Record
    {
        public string PriorYear { get; set; }
        public string CurrentYearByMonth { get; set; }
        public string YTDForecast { get; set; }
    }
}