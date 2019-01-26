using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.OPERATING_METRICS
{
    public class FinancialPerformance
    {
        public Record GR { get; set; }
        public Record NOI { get; set; }
        public Record OPM { get; set; }
        public Record EBITA { get; set; }
        public Record EBITPercent { get; set; }
        public Record TC { get; set; }
        public Record MB { get; set; }
        public Record NCF { get; set; }

        public FinancialPerformance()
        {
            GR = new Record();
            NOI = new Record();
            OPM = new Record();
            EBITA = new Record();
            EBITPercent = new Record();
            TC = new Record();
            MB = new Record();
            NCF = new Record();
        }
    }
}