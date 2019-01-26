using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models
{
    public class CustomerBase
    {
        public string CustomerBaseBy { get; set; }
        public Customers Customer1 { get; set; }
        public Customers Customer2 { get; set; }
        public Customers Customer3 { get; set; }
        public Customers Customer4 { get; set; }
        public Customers Customer5 { get; set; }
        public Customers Customer6 { get; set; }
        public Customers Customer7 { get; set; }
        public Customers Customer8 { get; set; }
        public Customers Customer9 { get; set; }
        public Customers Customer10 { get; set; }

        public CustomerBase()
        {
            CustomerBaseBy = "";
            Customer1 = new Customers();
            Customer2 = new Customers();
            Customer3 = new Customers();
            Customer4 = new Customers();
            Customer5 = new Customers();
            Customer6 = new Customers();
            Customer7 = new Customers();
            Customer8 = new Customers();
            Customer9 = new Customers();
            Customer10 = new Customers();
        }
    }

}