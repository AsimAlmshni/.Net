using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Quality
{
    public class ScrabInDollar
    {
        public decimal PriorYearTotal { get; set; }
        public decimal PriorYearMonthlyAverage { get; set; }
        public decimal CurrentYearMonthlyAverage { get; set; }
        public decimal YTDForecast { get; set; }
    }
}