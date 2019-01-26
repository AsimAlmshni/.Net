using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Productivity
{
    public class Productivity
    {
        public OperatingEfficiency OE { get; set; }

        public Productivity()
        {
            OE = new OperatingEfficiency();
        }
    }
}