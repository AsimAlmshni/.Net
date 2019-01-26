using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Quality2
{
    public class Quality2
    {
        public Record CAC { get; set; }
        public Record NOWI { get; set; }
        public Record NORP { get; set; }
        public Record NORRP { get; set; }
        public Record PPM { get; set; }
        public Record NOCITS { get; set; }
        public Record CBI { get; set; }
        public Record CBR { get; set; }
        public Record NOSIC { get; set; }
        public Record NOPRPIC { get; set; }

        public Quality2()
        {
            CAC = new Record();
            NOWI = new Record();
            NORP = new Record();
            NORRP = new Record();
            PPM = new Record();
            NOCITS = new Record();
            CBI = new Record();
            CBR = new Record();
            NOSIC = new Record();
            NOPRPIC = new Record();
        }
    }
}