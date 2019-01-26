using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Productivity
{
    public class Productivity2
    {
        public Record SP { get; set; }
        public Record SD { get; set; }
        public Record SPC { get; set; }
        public Record FTQ { get; set; }
        public Record OTP { get; set; }
        public Record OTH { get; set; }
        public Record OTD { get; set; }
        public Record DTP { get; set; }
        public Record DTH { get; set; }

        public Productivity2()
        {
            SP = new Record();
            SD = new Record();
            SPC = new Record();
            FTQ = new Record();
            OTP = new Record();
            OTH = new Record();
            OTD = new Record();
            DTP = new Record();
            DTH = new Record();
        }
    }
}