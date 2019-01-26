using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Inventory
{
    public class Record
    {
        public decimal PYT { get; set; }
        public decimal CYB { get; set; }
        public decimal CYMA { get; set; }
        public decimal YTDFocast { get; set; }
    }
}