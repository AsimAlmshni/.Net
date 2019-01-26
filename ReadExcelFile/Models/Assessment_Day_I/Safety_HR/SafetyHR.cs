using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Safety_HR
{
    public class SafetyHR
    {
        public Record NOR { get; set; }
        public Record NOLTWD { get; set; }

        public SafetyHR()
        {
            NOR = new Record();
            NOLTWD = new Record();
        }
    }
}