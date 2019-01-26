using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Cost
{
    public class Cost
    {
        public Record Sls { get; set; }
        public Record MCOS { get; set; }
        public Record LCOS { get; set; }
        public Record ESGA { get; set; }
        public Record MEOS { get; set; }
        public Record UC { get; set; }
        public Record T { get; set; }

        public Cost()
        {
            Sls = new Record();
            MCOS = new Record();
            LCOS = new Record();
            ESGA = new Record();
            MEOS = new Record();
            UC = new Record();
            T = new Record();
        }
    }
}