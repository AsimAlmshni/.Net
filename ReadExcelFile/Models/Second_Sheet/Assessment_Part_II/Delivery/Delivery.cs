using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Delivery
{
    public class Delivery
    {
        public Record COTDP { get; set; }
        public Record IOTRP { get; set; }
        public Record CPPD { get; set; }
        public Record SPPD { get; set; }
        public Record OPF { get; set; }
        public Record IPF { get; set; }

        public Delivery()
        {
            COTDP = new Record();
            IOTRP = new Record();
            CPPD = new Record();
            SPPD = new Record();
            OPF = new Record();
            IPF = new Record();
        }
    }
}