using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Quality
{
    public class Record
    {
        public string PriorYearTotal { get; set; }
        public string PriorYearMonthlyAverage { get; set; }
        public string CurrentYearMonthlyAverage { get; set; }
        public string YTDForecast { get; set; }
    }
}