using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models.Inventory
{
    public class Inventory
    {
        public Record COGSAYE { get; set; }
        public Record RIAYE { get; set; }
        public Record PIAYE { get; set; }
        public Record WIPAYE { get; set; }
        public Record FIAYE { get; set; }
        public Record TIAYE { get; set; }
        public TotalTurnsPerYear TTPY { get; set; }

        public Inventory()
        {
            COGSAYE = new Record();
            RIAYE = new Record();
            PIAYE = new Record();
            WIPAYE = new Record();
            FIAYE = new Record();
            TIAYE = new Record();
            TTPY = new TotalTurnsPerYear();
        }
    }
}