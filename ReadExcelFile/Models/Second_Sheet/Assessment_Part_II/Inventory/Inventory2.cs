using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Inventory
{
    public class Inventory2
    {
        public Record DOH { get; set; }
        public Record Raw { get; set; }
        public Record Purchased { get; set; }
        public Record WIP { get; set; }
        public Record Finished { get; set; }
        public Record DOH2 { get; set; }
        public Record Raw2 { get; set; }
        public Record Purchased2 { get; set; }
        public Record WIP2 { get; set; }
        public Record Finished2 { get; set; }

        public Inventory2()
        {
            DOH = new Record();
            Raw = new Record();
            Purchased = new Record();
            WIP = new Record();
            Finished = new Record();
            DOH2 = new Record();
            Raw2 = new Record();
            Purchased2 = new Record();
            WIP2 = new Record();
            Finished2 = new Record();
        }
    }
}