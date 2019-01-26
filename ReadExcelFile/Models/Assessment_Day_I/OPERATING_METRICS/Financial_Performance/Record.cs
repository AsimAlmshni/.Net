using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.OPERATING_METRICS
{
    public class Record
    {
        public decimal PriorYear { get; set; }
        public decimal CurrentYeadToDate { get; set; }
        public decimal CurrentBudget { get; set; }
    }
}