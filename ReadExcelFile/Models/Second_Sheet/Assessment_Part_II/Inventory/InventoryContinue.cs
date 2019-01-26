using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Inventory
{
    public class InventoryContinue
    {
        public Record IMOH { get; set; }
        public Record Raw { get; set; }
        public Record Purchased { get; set; }
        public Record WIP { get; set; }
        public Record Finished { get; set; }
        public Record TIM { get; set; }
        public Record ExOb { get; set; }

        public InventoryContinue()
        {
            IMOH = new Record();
            Raw = new Record();
            Purchased = new Record();
            WIP = new Record();
            Finished = new Record();
            TIM = new Record();
            ExOb = new Record();
        }
    }
}