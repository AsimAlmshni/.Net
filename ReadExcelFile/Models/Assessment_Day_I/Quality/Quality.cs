using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Quality
{
    public class Quality
    {
        public Record NOCC { get; set; }
        public Record QREC { get; set; }
        public Record QRIO { get; set; }
        public ScrabInDollar SID { get; set; }

        public Quality()
        {
            NOCC = new Record();
            QREC = new Record();
            QRIO = new Record();
            SID = new ScrabInDollar();
        }
    }
}