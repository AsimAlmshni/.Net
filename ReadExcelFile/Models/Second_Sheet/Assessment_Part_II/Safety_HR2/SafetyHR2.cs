using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Safety_HR2
{
    public class SafetyHR2
    {
        public Record ED { get; set; }
        public Record AA { get; set; }
        public Record G { get; set; }
        public Record ETO { get; set; }
        public Record NOG { get; set; }
        public Record EI { get; set; }

        public SafetyHR2()
        {
            ED = new Record();
            AA = new Record();
            G = new Record();
            ETO = new Record();
            NOG = new Record();
            EI = new Record();
        }
    }
}