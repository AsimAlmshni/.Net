using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelFile.Models
{
    public class TypeOfProcesses
    {
        public string MetalCutting { get; set; }
        public string MetalFormingSheet { get; set; }
        public string MetalFormingForging { get; set; }
        public string AssembleyTest { get; set; }
        public string PaintingCoating { get; set; }
        public string MetalFabrication { get; set; }
        public string Casting { get; set; }
        public string Welding { get; set; }
    }
}