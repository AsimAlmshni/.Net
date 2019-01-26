using log4net;
using log4net.Config;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using ReadExcelFile.Models;
using ReadExcelFile.Models.Inventory;
using ReadExcelFile.Models.Notes_And_Comments.CLEAN_ORDERLY_ORGANIZED;
using ReadExcelFile.Models.OPERATING_METRICS;
using ReadExcelFile.Models.Productivity;
using ReadExcelFile.Models.Quality;
using ReadExcelFile.Models.Safety_HR;
using ReadExcelFile.Models.Second_Sheet;
using ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Cost;
using ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Delivery;
using ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Inventory;
using ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Productivity;
using ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Quality2;
using ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Safety_HR2;
using ReadExcelFile.Models.Second_Sheet.Assessment_Part_II.Standards;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

//[assembly: log4net.Config.XmlConfigurator(ConfigFile = "Web.config", Watch = true)]


namespace ReadExcelFile.Controllers
{
    public class HomeController : Controller
    {
        private static log4net.ILog Log { get; set; }
        //ILog log = log4net.LogManager.GetLogger(typeof(HomeController));
        private static readonly ILog log = LogManager.GetLogger(typeof(HomeController));



        // Objects for first sheet
        private List<object> opj = new List<object>();
        private List<Employees> Emp = new List<Employees>();
        private List<CompanyLocation> CoLoc = new List<CompanyLocation>();
        private List<TypeOfProcesses> TOP = new List<TypeOfProcesses>();
        private List<CustomerBase> CB = new List<CustomerBase>();
        private List<Products> P = new List<Products>();
        private List<FinancialPerformance> FP = new List<FinancialPerformance>();
        private List<SafetyHR> SHR = new List<SafetyHR>();
        private List<Quality> Q = new List<Quality>();
        private List<Productivity> Productive = new List<Productivity>();
        private List<Inventory> Inv = new List<Inventory>();
        private List<Customer> CustomerDetails = new List<Customer>();

        // Objects for second sheet
        private List<CompanyDetails> CoDetails = new List<CompanyDetails>();
        private List<Standards> Std = new List<Standards>();
        private List<SafetyHR2> SHR2 = new List<SafetyHR2>();
        private List<Quality2> Q2 = new List<Quality2>();
        private List<Delivery> Dl = new List<Delivery>();
        private List<Productivity2> P2 = new List<Productivity2>();
        private List<Inventory2> Inv2 = new List<Inventory2>();
        private List<InventoryContinue> Inv2Con = new List<InventoryContinue>();
        private List<Cost> Cst = new List<Cost>();

        // Objects for third sheet
        private List<CompanyDetails> CoDetails3 = new List<CompanyDetails>();
        private List<CleanOrderlyOrganized> COO = new List<CleanOrderlyOrganized>();
        private List<CleanOrderlyOrganized> COO2 = new List<CleanOrderlyOrganized>();
        private List<CleanOrderlyOrganized> COO3 = new List<CleanOrderlyOrganized>();
        private List<CleanOrderlyOrganized> COO4 = new List<CleanOrderlyOrganized>();
        private List<CleanOrderlyOrganized> COO5 = new List<CleanOrderlyOrganized>();
        private List<CleanOrderlyOrganized> COO6 = new List<CleanOrderlyOrganized>();
        private List<CleanOrderlyOrganized> COO7 = new List<CleanOrderlyOrganized>();

        private int RowIdx;
        private string ColIdx; 

        //log4net.ILog logger = log4net.LogManager.GetLogger(typeof(HomeController));  //Declaring Log4Net  

        public ActionResult Index()
        {
            //BasicConfigurator.Configure();

            XmlConfigurator.Configure(new FileInfo("log4net.config"));
            
            //log.Debug("This is a Debug message");

            //log.Info("This is a Info message");

            //log.Warn("This is a Warning message");

            //log.Error("This is an Error message");

            //log.Fatal("This is a Fatal message");

            ReadExcelFile();
            // First Sheet
            ViewBag.pront = CustomerDetails;
            ViewBag.printLoc = CoLoc;
            ViewBag.EmployeeData = Emp;
            ViewBag.TP = TOP;
            ViewBag.CBase = CB;
            ViewBag.Product = P;
            ViewBag.F = FP;
            ViewBag.HR = SHR;
            ViewBag.PQuality = Q;
            ViewBag.PProductivitiy = Productive;
            ViewBag.PInventory = Inv;

            // Second Sheet
            ViewBag.SecondSheet = CoDetails;
            ViewBag.STD = Std;
            ViewBag.SH2 = SHR2;
            ViewBag.QQ2 = Q2;
            ViewBag.Dlv = Dl;
            ViewBag.Pr2 = P2;
            ViewBag.Iv2 = Inv2;
            ViewBag.InvCon = Inv2Con;
            ViewBag.CostData = Cst;

            // Third Sheet
            ViewBag.CompDetails3 = CoDetails3;
            ViewBag.Coo = COO;
            ViewBag.Coo2 = COO2;
            ViewBag.Coo3 = COO3;
            ViewBag.Coo4 = COO4;
            ViewBag.Coo5 = COO5;
            ViewBag.Coo6 = COO6;
            ViewBag.Coo7 = COO7;

            return View();
        }

        public void ReadExcelFile()
        {
            //Excel.Application xlApp = new Excel.Application();
            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\aasim\Desktop\Emplyees.xlsx");
            //Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            //Excel.Range xlRange = xlWorksheet.UsedRange;

            //int rowCount = xlRange.Rows.Count;
            //int colCount = xlRange.Columns.Count;

            //for (int row = 2; row <= xlRange.Rows.Count; row++)
            //{
            //    Emp.Add( new Employee{
            //        Email = ((Excel.Range)xlRange.Cells[row, 1]).Text,
            //        FirstName = ((Excel.Range)xlRange.Cells[row, 2]).Text,
            //        LastName = ((Excel.Range)xlRange.Cells[row, 3]).Text,
            //        Salary = int.Parse(((Excel.Range)xlRange.Cells[row, 4]).Text),
            //        Age = int.Parse(((Excel.Range)xlRange.Cells[row, 5]).Text)
            //});
            //}

            ExcelWorksheet currentWorksheet = null;
            
            //Reading more than single tab in single Excel File
            var file = new FileInfo(@"C:\Users\aasim\Desktop\HVMC_Sheet.xlsx");
            using (var package = new ExcelPackage(file))
            {
                    // Get the work book in the file
                ExcelWorkbook workBook = package.Workbook; //Hangs here for about 2 mins
                int pageCount = 0;
                if (workBook != null)
                {
                    foreach (var page in workBook.Worksheets) {

                        currentWorksheet = page;

                        if (currentWorksheet.Name.Equals("Assessment Day 1"))
                        {
                            Customer tempCustomer = new Customer();
                            CompanyLocation tempLocation = new CompanyLocation();
                            Employees tempEmp = new Employees();
                            TypeOfProcesses tempTOP = new TypeOfProcesses();
                            CustomerBase tempCB = new CustomerBase();
                            Customers Cs = new Customers();
                            Products p = new Products();
                            FinancialPerformance fp = new FinancialPerformance();
                            SafetyHR shr = new SafetyHR();
                            Quality q = new Quality();
                            Productivity tempProductivity = new Productivity();
                            Inventory tempInventory = new Inventory();

                            //  get info about company name and completed by and date complete 
                            for (int i = 2; i <= 4; i++)
                            {
                                try
                                {
                                    var name = currentWorksheet.Cells[i, 9].Value;

                                    RowIdx = i;
                                    ColIdx = "I";

                                    if (i == 2)
                                    {
                                        tempCustomer.CompanyName = name.ToString();
                                    }
                                    else if (i == 3)
                                    {
                                        tempCustomer.CompletedBy = name.ToString();
                                    }
                                    else if (i == 4)
                                    {
                                        tempCustomer.DateCompleted = name.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx  +" /n The Exception is /n" + ex.ToString());
                                }
                            }
                            CustomerDetails.Add(tempCustomer);

                            //  get info about company Location and other properties
                            for (int i = 6; i <= 8; i++)
                            {
                                try { 
                                    var temp = currentWorksheet.Cells[i, 2].Value;

                                    RowIdx = i;
                                    ColIdx = "I";

                                    if (i == 6)
                                    {
                                        tempLocation.Location = temp.ToString();
                                    }
                                    else if (i == 7)
                                    {
                                        tempLocation.SquareFootage = temp.ToString();
                                    }
                                    else if (i == 8)
                                    {
                                        tempLocation.Union = temp.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }

                        }
                        CoLoc.Add(tempLocation);


                            //  get info about Employees/Workers properties
                            for (int i = 7; i <= 11; i++)
                            {
                                try {
                                    var temp = currentWorksheet.Cells["F" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "F";

                                    if (i == 7)
                                    {
                                        tempEmp.NoOfSalariedEmployees = temp.ToString();
                                    }
                                    else if (i == 8)
                                    {
                                        tempEmp.NoOfHourlyEmployee = temp.ToString();
                                    }
                                    else if (i == 9)
                                    {
                                        tempEmp.NoOfTemprorayWorkers = temp.ToString();
                                    }
                                    else if (i == 10)
                                    {
                                        tempEmp.NoOfSkilledEmployees = temp.ToString();
                                    }
                                    else if (i == 11)
                                    {
                                        tempEmp.ActualTotalNoOfEmplyees = int.Parse(temp.ToString());
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }

                            }
                            Emp.Add(tempEmp);

                            //  get info about Type Of Processes
                            for (int i = 7; i <= 14; i++)
                            {
                                try {
                                    var temp = currentWorksheet.Cells["I" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "I";

                                    if (i == 7)
                                    {
                                        tempTOP.MetalCutting = temp.ToString();
                                    }
                                    else if (i == 8)
                                    {
                                        tempTOP.MetalFormingSheet = temp.ToString();
                                    }
                                    else if (i == 9)
                                    {
                                        tempTOP.MetalFormingForging = temp.ToString();
                                    }
                                    else if (i == 10)
                                    {
                                        tempTOP.AssembleyTest = temp.ToString();
                                    }
                                    else if (i == 11)
                                    {
                                        tempTOP.PaintingCoating = temp.ToString();
                                    }
                                    else if (i == 12)
                                    {
                                        tempTOP.MetalFabrication = temp.ToString();
                                    }
                                    else if (i == 13)
                                    {
                                        tempTOP.Casting = temp.ToString();
                                    }
                                    else if (i == 14)
                                    {
                                        tempTOP.Welding = temp.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }

                            }
                            TOP.Add(tempTOP);

                            //  get info about Customer Base
                            for (int i = 6; i <= 17; i++)
                            {
                                try {
                                    var temp = currentWorksheet.Cells["L" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["M" + i + ""].Value;

                                    RowIdx = i;

                                    if (i == 6)
                                    {
                                        ViewBag.head = temp.ToString();
                                        continue;
                                    }
                                    else if (i == 8)
                                    {
                                        ColIdx = "L";
                                        tempCB.Customer1.Revenue = decimal.Parse(temp.ToString());

                                        ColIdx = "M";
                                        tempCB.Customer1.TotalRevenue = decimal.Parse(temp2.ToString());
                                    }
                                    else if (i == 9)
                                    {
                                        ColIdx = "L";
                                        tempCB.Customer2.Revenue = decimal.Parse(temp.ToString());
                                        ColIdx = "M";
                                        tempCB.Customer2.TotalRevenue = decimal.Parse(temp2.ToString()); ;
                                    }
                                    else if (i == 10)
                                    {
                                        ColIdx = "L";
                                        tempCB.Customer3.Revenue = decimal.Parse(temp.ToString());
                                        ColIdx = "M";
                                        tempCB.Customer3.TotalRevenue = decimal.Parse(temp2.ToString());
                                    }
                                    else if (i == 11)
                                    {
                                        ColIdx = "L";
                                        tempCB.Customer4.Revenue = decimal.Parse(temp.ToString());
                                        ColIdx = "M";
                                        tempCB.Customer4.TotalRevenue = decimal.Parse(temp2.ToString());
                                    }
                                    else if (i == 12)
                                    {
                                        ColIdx = "L";
                                        tempCB.Customer5.Revenue = decimal.Parse(temp.ToString());
                                        ColIdx = "M";
                                        tempCB.Customer5.TotalRevenue = decimal.Parse(temp2.ToString());
                                    }
                                    else if (i == 13)
                                    {
                                        ColIdx = "L";
                                        tempCB.Customer6.Revenue = decimal.Parse(temp.ToString());
                                        ColIdx = "M";
                                        tempCB.Customer6.TotalRevenue = decimal.Parse(temp2.ToString());
                                    }
                                    else if (i == 14)
                                    {
                                        ColIdx = "L";
                                        tempCB.Customer7.Revenue = decimal.Parse(temp.ToString());
                                        ColIdx = "M";
                                        tempCB.Customer7.TotalRevenue = decimal.Parse(temp2.ToString());
                                    }
                                    else if (i == 15)
                                    {
                                        ColIdx = "L";
                                        tempCB.Customer8.Revenue = decimal.Parse(temp.ToString());
                                        ColIdx = "M";
                                        tempCB.Customer8.TotalRevenue = decimal.Parse(temp2.ToString());
                                    }
                                    else if (i == 16)
                                    {
                                        ColIdx = "L";
                                        tempCB.Customer9.Revenue = decimal.Parse(temp.ToString());
                                        ColIdx = "M";
                                        tempCB.Customer9.TotalRevenue = decimal.Parse(temp2.ToString());
                                    }
                                    else if (i == 17)
                                    {
                                        ColIdx = "L";
                                        tempCB.Customer10.Revenue = decimal.Parse(temp.ToString());
                                        ColIdx = "M";
                                        tempCB.Customer10.TotalRevenue = decimal.Parse(temp2.ToString());
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }

                            }
                            CB.Add(tempCB);

                            //  get info about Products properties
                            for (int i = 11; i <= 12; i++)
                            {
                                try {
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "B";

                                    if (i == 11)
                                    {
                                        p.MajorProductLine = decimal.Parse(temp.ToString());
                                    }
                                    else if (i == 12)
                                    {
                                        p.NumberOfSuppliers = decimal.Parse(temp.ToString());
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            P.Add(p);

                            //  get info about Financial Performance
                            for (int i = 17; i <= 24; i++)
                            {
                                try { 
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["D" + i + ""].Value;
                                    var temp3 = currentWorksheet.Cells["F" + i + ""].Value;

                                    RowIdx = i;

                                    if (i == 17)
                                    {
                                        ColIdx = "B";
                                        fp.GR.PriorYear = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        fp.GR.CurrentYeadToDate = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        fp.GR.CurrentBudget = decimal.Parse(temp3.ToString());
                                    }
                                    else if (i == 18)
                                    {
                                        ColIdx = "B";
                                        fp.NOI.PriorYear = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        fp.NOI.CurrentYeadToDate = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        fp.NOI.CurrentBudget = decimal.Parse(temp3.ToString());
                                    }
                                    else if (i == 19)
                                    {
                                        ColIdx = "B";
                                        fp.OPM.PriorYear = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        fp.OPM.CurrentYeadToDate = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        fp.OPM.CurrentBudget = decimal.Parse(temp3.ToString());
                                    }
                                    else if (i == 20)
                                    {
                                        ColIdx = "B";
                                        fp.EBITA.PriorYear = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        fp.EBITA.CurrentYeadToDate = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        fp.EBITA.CurrentBudget = decimal.Parse(temp3.ToString());
                                    }
                                    else if (i == 21)
                                    {
                                        ColIdx = "B";
                                        fp.EBITPercent.PriorYear = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        fp.EBITPercent.CurrentYeadToDate = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        fp.EBITPercent.CurrentBudget = decimal.Parse(temp3.ToString());
                                    }
                                    else if (i == 22)
                                    {
                                        ColIdx = "B";
                                        fp.TC.PriorYear = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        fp.TC.CurrentYeadToDate = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        fp.TC.CurrentBudget = decimal.Parse(temp3.ToString());
                                    }
                                    else if (i == 23)
                                    {
                                        ColIdx = "B";
                                        fp.MB.PriorYear = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        fp.MB.CurrentYeadToDate = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        fp.MB.CurrentBudget = decimal.Parse(temp3.ToString());
                                    }
                                    else if (i == 24)
                                    {
                                        ColIdx = "B";
                                        fp.NCF.PriorYear = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        fp.NCF.CurrentYeadToDate = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        fp.NCF.CurrentBudget = decimal.Parse(temp3.ToString());
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }

                            }
                            FP.Add(fp);


                            //  get info about Safety/HR
                            for (int i = 27; i <= 28; i++)
                            {
                                try { 
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["D" + i + ""].Value;
                                    var temp3 = currentWorksheet.Cells["F" + i + ""].Value;
                                    var temp4 = currentWorksheet.Cells["H" + i + ""].Value;

                                    RowIdx = i;

                                    if (i == 27)
                                    {
                                        ColIdx = "B";
                                        shr.NOR.PriorYearTotal = temp.ToString();
                                        ColIdx = "D";
                                        shr.NOR.PYMA = temp2.ToString();
                                        ColIdx = "F";
                                        shr.NOR.CYMA = temp3.ToString();
                                        ColIdx = "H";
                                        shr.NOR.YTDForecast = temp4.ToString();
                                    }
                                    else if (i == 28)
                                    {
                                        ColIdx = "B";
                                        shr.NOLTWD.PriorYearTotal = temp.ToString();
                                        ColIdx = "D";
                                        shr.NOLTWD.PYMA = temp2.ToString();
                                        ColIdx = "F";
                                        shr.NOLTWD.CYMA = temp3.ToString();
                                        ColIdx = "H";
                                        shr.NOLTWD.YTDForecast = temp4.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            SHR.Add(shr);

                            //  get info about Quality 
                            for (int i = 31; i <= 34; i++)
                            {
                                try { 
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["D" + i + ""].Value;
                                    var temp3 = currentWorksheet.Cells["F" + i + ""].Value;
                                    var temp4 = currentWorksheet.Cells["H" + i + ""].Value;

                                    RowIdx = i;
                                    if (i == 31)
                                    {
                                        ColIdx = "B";
                                        q.NOCC.PriorYearTotal = temp.ToString();
                                        ColIdx = "D";
                                        q.NOCC.PriorYearMonthlyAverage = temp2.ToString();
                                        ColIdx = "F";
                                        q.NOCC.CurrentYearMonthlyAverage = temp3.ToString();
                                        ColIdx = "H";
                                        q.NOCC.YTDForecast = temp4.ToString();
                                    }
                                    else if (i == 32)
                                    {
                                        ColIdx = "B";
                                        q.QREC.PriorYearTotal = temp.ToString();
                                        ColIdx = "D";
                                        q.QREC.PriorYearMonthlyAverage = temp2.ToString();
                                        ColIdx = "F";
                                        q.QREC.CurrentYearMonthlyAverage = temp3.ToString();
                                        ColIdx = "H";
                                        q.QREC.YTDForecast = temp4.ToString();
                                    }
                                    else if (i == 33)
                                    {
                                        ColIdx = "B";
                                        q.QRIO.PriorYearTotal = temp.ToString();
                                        ColIdx = "D";
                                        q.QRIO.PriorYearMonthlyAverage = temp2.ToString();
                                        ColIdx = "F";
                                        q.QRIO.CurrentYearMonthlyAverage = temp3.ToString();
                                        ColIdx = "H";
                                        q.QRIO.YTDForecast = temp4.ToString();
                                    }
                                    else if (i == 34)
                                    {
                                        ColIdx = "B";
                                        q.SID.PriorYearTotal = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        q.SID.PriorYearMonthlyAverage = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        q.SID.CurrentYearMonthlyAverage = decimal.Parse(temp3.ToString());
                                        ColIdx = "H";
                                        q.SID.YTDForecast = decimal.Parse(temp4.ToString());
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            Q.Add(q);

                            //  get info about Productivity 
                            for (int i = 37; i <= 37; i++)
                            {
                                try {
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["D" + i + ""].Value;
                                    var temp3 = currentWorksheet.Cells["F" + i + ""].Value;
                                    var temp4 = currentWorksheet.Cells["H" + i + ""].Value;

                                    RowIdx = i;

                                    if (i == 37)
                                    {
                                        ColIdx = "B";
                                        tempProductivity.OE.PYT = temp.ToString();
                                        ColIdx = "D";
                                        tempProductivity.OE.PYMA = temp2.ToString();
                                        ColIdx = "F";
                                        tempProductivity.OE.CYMA = temp3.ToString();
                                        ColIdx = "H";
                                        tempProductivity.OE.PTDForcats = temp4.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            Productive.Add(tempProductivity);

                            //  get info about Inventory 
                            for (int i = 41; i <= 49; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["D" + i + ""].Value;
                                    var temp3 = currentWorksheet.Cells["F" + i + ""].Value;
                                    var temp4 = currentWorksheet.Cells["H" + i + ""].Value;

                                    RowIdx = i;

                                    if (i == 41)
                                    {
                                        ColIdx = "B";
                                        tempInventory.COGSAYE.PYT = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        tempInventory.COGSAYE.CYB = decimal.Parse(temp.ToString());
                                        ColIdx = "F";
                                        tempInventory.COGSAYE.CYMA = decimal.Parse(temp.ToString());
                                        ColIdx = "H";
                                        tempInventory.COGSAYE.YTDFocast = decimal.Parse(temp.ToString());
                                        continue;
                                    }
                                    else if (i == 43)
                                    {
                                        ColIdx = "B";
                                        tempInventory.RIAYE.PYT = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        tempInventory.RIAYE.CYB = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        tempInventory.RIAYE.CYMA = decimal.Parse(temp3.ToString());
                                        ColIdx = "H";
                                        tempInventory.RIAYE.YTDFocast = decimal.Parse(temp4.ToString());
                                    }
                                    else if (i == 44)
                                    {
                                        ColIdx = "B";
                                        tempInventory.PIAYE.PYT = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        tempInventory.PIAYE.CYB = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        tempInventory.PIAYE.CYMA = decimal.Parse(temp3.ToString());
                                        ColIdx = "H";
                                        tempInventory.PIAYE.YTDFocast = decimal.Parse(temp4.ToString());
                                    }
                                    else if (i == 45)
                                    {
                                        ColIdx = "B";
                                        tempInventory.WIPAYE.PYT = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        tempInventory.WIPAYE.CYB = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        tempInventory.WIPAYE.CYMA = decimal.Parse(temp3.ToString());
                                        ColIdx = "H";
                                        tempInventory.WIPAYE.YTDFocast = decimal.Parse(temp4.ToString());
                                    }
                                    else if (i == 46)
                                    {
                                        ColIdx = "B";
                                        tempInventory.FIAYE.PYT = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        tempInventory.FIAYE.CYB = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        tempInventory.FIAYE.CYMA = decimal.Parse(temp3.ToString());
                                        ColIdx = "H";
                                        tempInventory.FIAYE.YTDFocast = decimal.Parse(temp4.ToString());
                                    }
                                    else if (i == 47)
                                    {
                                        ColIdx = "B";
                                        tempInventory.TIAYE.PYT = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        tempInventory.TIAYE.CYB = decimal.Parse(temp2.ToString());
                                        ColIdx = "F";
                                        tempInventory.TIAYE.CYMA = decimal.Parse(temp3.ToString());
                                        ColIdx = "H";
                                        tempInventory.TIAYE.YTDFocast = decimal.Parse(temp4.ToString());
                                        continue;
                                    }
                                    else if (i == 49)
                                    {
                                        ColIdx = "B";
                                        tempInventory.TTPY.PYT = decimal.Parse(temp.ToString());
                                        ColIdx = "D";
                                        tempInventory.TTPY.CYB = temp2.ToString();
                                        ColIdx = "F";
                                        tempInventory.TTPY.CYMA = temp3.ToString();
                                        ColIdx = "H";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            Inv.Add(tempInventory);
                            pageCount++;
                        }
                        // Second Excel Sheet
                        else if (currentWorksheet.Name.Equals("Assessment Day 2"))
                        {
                            // get the next sheet
                            CompanyDetails cD = new CompanyDetails();
                            Standards std = new Standards();
                            SafetyHR2 shr2 = new SafetyHR2();
                            Quality2 q = new Quality2();
                            Delivery d = new Delivery();
                            Productivity2 p2 = new Productivity2();
                            Inventory2 inv2 = new Inventory2();
                            InventoryContinue invCon = new InventoryContinue();
                            Cost cst = new Cost();

                            // get info about company like name and date
                            for (int i = 2; i <= 4; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["F" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "F";
                                    if (i == 2)
                                    {
                                        cD.CompanyName = temp.ToString();
                                    }
                                    else if (i == 3)
                                    {
                                        cD.CompletedBy = temp.ToString();
                                    }
                                    else if (i == 4)
                                    {
                                        cD.Date = temp.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }

                            }
                            CoDetails.Add(cD);

                            // get info about Standards
                            for (int i = 7; i <= 10; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "B";

                                    if (i == 7)
                                    {
                                        std.LaborStandards = temp.ToString();
                                    }
                                    else if (i == 8)
                                    {
                                        std.MaterialStandards = temp.ToString();
                                    }
                                    else if (i == 9)
                                    {
                                        std.CycleTimeStandards = temp.ToString();
                                    }
                                    else if (i == 10)
                                    {
                                        std.ProductionPcsStandards = temp.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            Std.Add(std);

                            // get info about SafteyHR in 2nd sheet
                            for (int i = 13; i <= 19; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["D" + i + ""].Value;
                                    var temp3 = currentWorksheet.Cells["F" + i + ""].Value;

                                    RowIdx = i;

                                    if (i == 13)
                                    {
                                        ColIdx = "B";
                                        shr2.ED.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        shr2.ED.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        shr2.ED.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 14)
                                    {
                                        ColIdx = "B";
                                        shr2.AA.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        shr2.AA.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        shr2.AA.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 15)
                                    {
                                        ColIdx = "B";
                                        shr2.G.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        shr2.G.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        shr2.G.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 16)
                                    {
                                        ColIdx = "B";
                                        shr2.ETO.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        shr2.ETO.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        shr2.ETO.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 18)
                                    {
                                        ColIdx = "B";
                                        shr2.NOG.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        shr2.NOG.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        shr2.NOG.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 19)
                                    {
                                        ColIdx = "B";
                                        shr2.EI.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        shr2.EI.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        shr2.EI.YTDForecast = temp3.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            SHR2.Add(shr2);

                            // get info about Quality in 2nd sheet
                            for (int i = 23; i <= 33; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["D" + i + ""].Value;
                                    var temp3 = currentWorksheet.Cells["F" + i + ""].Value;

                                    RowIdx = i;
                                    if (i == 23)
                                    {
                                        ColIdx = "B";
                                        q.CAC.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        q.CAC.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        q.CAC.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 24)
                                    {
                                        ColIdx = "B";
                                        q.NOWI.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        q.NOWI.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        q.NOWI.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 25)
                                    {
                                        ColIdx = "B";
                                        q.NORP.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        q.NORP.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        q.NORP.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 26)
                                    {
                                        ColIdx = "B";
                                        q.NORRP.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        q.NORRP.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        q.NORRP.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 28)
                                    {
                                        ColIdx = "B";
                                        q.PPM.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        q.PPM.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        q.PPM.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 29)
                                    {
                                        ColIdx = "B";
                                        q.NOCITS.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        q.NOCITS.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        q.NOCITS.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 30)
                                    {
                                        ColIdx = "B";
                                        q.CBI.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        q.CBI.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        q.CBI.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 31)
                                    {
                                        ColIdx = "B";
                                        q.CBR.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        q.CBR.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        q.CBR.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 32)
                                    {
                                        ColIdx = "B";
                                        q.NOSIC.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        q.NOSIC.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        q.NOSIC.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 33)
                                    {
                                        ColIdx = "B";
                                        q.NOPRPIC.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        q.NOPRPIC.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        q.NOPRPIC.YTDForecast = temp3.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            Q2.Add(q);


                            // get info about Quality in 2nd sheet
                            for (int i = 37; i <= 42; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["D" + i + ""].Value;
                                    var temp3 = currentWorksheet.Cells["F" + i + ""].Value;

                                    RowIdx = i;
                                    if (i == 37)
                                    {
                                        ColIdx = "B";
                                        d.COTDP.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        d.COTDP.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        d.COTDP.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 38)
                                    {
                                        ColIdx = "B";
                                        d.IOTRP.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        d.IOTRP.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        d.IOTRP.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 39)
                                    {
                                        ColIdx = "B";
                                        d.CPPD.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        d.CPPD.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        d.CPPD.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 40)
                                    {
                                        ColIdx = "B";
                                        d.SPPD.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        d.SPPD.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        d.SPPD.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 41)
                                    {
                                        ColIdx = "B";
                                        d.OPF.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        d.OPF.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        d.OPF.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 42)
                                    {
                                        ColIdx = "B";
                                        d.IPF.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        d.IPF.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        d.IPF.YTDForecast = temp3.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            Dl.Add(d);

                            // get info about Productivity in 2nd sheet
                            for (int i = 46; i <= 54; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["D" + i + ""].Value;
                                    var temp3 = currentWorksheet.Cells["F" + i + ""].Value;

                                    RowIdx = i;

                                    if (i == 46)
                                    {
                                        ColIdx = "B";
                                        p2.SP.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        p2.SP.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        p2.SP.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 47)
                                    {
                                        ColIdx = "B";
                                        p2.SD.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        p2.SD.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        p2.SD.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 48)
                                    {
                                        ColIdx = "B";
                                        p2.SPC.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        p2.SPC.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        p2.SPC.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 49)
                                    {
                                        ColIdx = "B";
                                        p2.FTQ.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        p2.FTQ.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        p2.FTQ.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 50)
                                    {
                                        ColIdx = "B";
                                        p2.OTP.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        p2.OTP.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        p2.OTP.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 51)
                                    {
                                        ColIdx = "B";
                                        p2.OTH.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        p2.OTH.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        p2.OTH.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 52)
                                    {
                                        ColIdx = "B";
                                        p2.OTD.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        p2.OTD.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        p2.OTD.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 53)
                                    {
                                        ColIdx = "B";
                                        p2.DTP.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        p2.DTP.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        p2.DTP.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 54)
                                    {
                                        ColIdx = "B";
                                        p2.DTH.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        p2.DTH.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        p2.DTH.YTDForecast = temp3.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            P2.Add(p2);

                            // get info about Inventory in 2nd sheet
                            for (int i = 58; i <= 67; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["D" + i + ""].Value;
                                    var temp3 = currentWorksheet.Cells["F" + i + ""].Value;

                                    RowIdx = i;

                                    if (i == 58)
                                    {
                                        ColIdx = "B";
                                        inv2.DOH.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        inv2.DOH.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        inv2.DOH.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 59)
                                    {
                                        ColIdx = "B";
                                        inv2.Raw.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        inv2.Raw.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        inv2.Raw.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 60)
                                    {
                                        ColIdx = "B";
                                        inv2.Purchased.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        inv2.Purchased.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        inv2.Purchased.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 61)
                                    {
                                        ColIdx = "B";
                                        inv2.WIP.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        inv2.WIP.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        inv2.WIP.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 62)
                                    {
                                        ColIdx = "B";
                                        inv2.Finished.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        inv2.Finished.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        inv2.Finished.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 63)
                                    {
                                        ColIdx = "B";
                                        inv2.DOH2.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        inv2.DOH2.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        inv2.DOH2.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 64)
                                    {
                                        ColIdx = "B";
                                        inv2.Raw2.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        inv2.Raw2.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        inv2.Raw2.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 65)
                                    {
                                        ColIdx = "B";
                                        inv2.Purchased2.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        inv2.Purchased2.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        inv2.Purchased2.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 66)
                                    {
                                        ColIdx = "B";
                                        inv2.WIP2.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        inv2.WIP2.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        inv2.WIP2.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 67)
                                    {
                                        ColIdx = "B";
                                        inv2.Finished2.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        inv2.Finished2.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        inv2.Finished2.YTDForecast = temp3.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            Inv2.Add(inv2);

                            // get info about Inventory Continued in 2nd sheet
                            for (int i = 70; i <= 76; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["D" + i + ""].Value;
                                    var temp3 = currentWorksheet.Cells["F" + i + ""].Value;

                                    RowIdx = i;

                                    if (i == 70)
                                    {
                                        ColIdx = "B";
                                        invCon.IMOH.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        invCon.IMOH.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        invCon.IMOH.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 71)
                                    {
                                        ColIdx = "B";
                                        invCon.Raw.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        invCon.Raw.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        invCon.Raw.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 72)
                                    {
                                        ColIdx = "B";
                                        invCon.Purchased.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        invCon.Purchased.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        invCon.Purchased.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 73)
                                    {
                                        ColIdx = "B";
                                        invCon.WIP.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        invCon.WIP.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        invCon.WIP.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 74)
                                    {
                                        ColIdx = "B";
                                        invCon.Finished.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        invCon.Finished.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        invCon.Finished.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 75)
                                    {
                                        ColIdx = "B";
                                        invCon.TIM.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        invCon.TIM.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        invCon.TIM.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 76)
                                    {
                                        ColIdx = "B";
                                        invCon.ExOb.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        invCon.ExOb.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        invCon.ExOb.YTDForecast = temp3.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            Inv2Con.Add(invCon);

                            // get info about Cost information in 2nd sheet
                            for (int i = 80; i <= 86; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["B" + i + ""].Value;
                                    var temp2 = currentWorksheet.Cells["D" + i + ""].Value;
                                    var temp3 = currentWorksheet.Cells["F" + i + ""].Value;

                                    RowIdx = i;

                                    if (i == 80)
                                    {
                                        ColIdx = "B";
                                        cst.Sls.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        cst.Sls.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        cst.Sls.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 81)
                                    {
                                        ColIdx = "B";
                                        cst.MCOS.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        cst.MCOS.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        cst.MCOS.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 82)
                                    {
                                        ColIdx = "B";
                                        cst.LCOS.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        cst.LCOS.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        cst.LCOS.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 83)
                                    {
                                        ColIdx = "B";
                                        cst.ESGA.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        cst.ESGA.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        cst.ESGA.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 84)
                                    {
                                        ColIdx = "B";
                                        cst.MEOS.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        cst.MEOS.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        cst.MEOS.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 85)
                                    {
                                        ColIdx = "B";
                                        cst.UC.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        cst.UC.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        cst.UC.YTDForecast = temp3.ToString();
                                    }
                                    else if (i == 86)
                                    {
                                        ColIdx = "B";
                                        cst.T.PriorYear = temp.ToString();
                                        ColIdx = "D";
                                        cst.T.CurrentYearByMonth = temp2.ToString();
                                        ColIdx = "F";
                                        cst.T.YTDForecast = temp3.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            Cst.Add(cst);

                            pageCount++;
                        }
                        else if (currentWorksheet.Name.Equals("Notes & Comments"))
                        {
                            CleanOrderlyOrganized cOO = new CleanOrderlyOrganized();

                            CompanyDetails cD = new CompanyDetails();
                            // get info about company like name and date
                            for (int i = 2; i <= 4; i++)
                            {
                                try
                                {
                                    var temp = currentWorksheet.Cells["C" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "C";

                                    if (i == 2)
                                    {
                                        cD.CompanyName = temp.ToString();
                                    }
                                    else if (i == 3)
                                    {
                                        cD.CompletedBy = temp.ToString();
                                    }
                                    else if (i == 4)
                                    {
                                        cD.Date = temp.ToString();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            CoDetails3.Add(cD);

                            for (int i = 6; i <= 13; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["C" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "C";

                                    if (i == 6)
                                    {
                                        cOO.A1 = temp.ToString();
                                    }
                                    else if (i == 8)
                                    {
                                        cOO.A2 = temp.ToString();
                                    }
                                    else if (i == 10)
                                    {
                                        cOO.A3 = temp.ToString();
                                    }
                                    else if (i == 13)
                                    {
                                        cOO.Observation = GetCellValueFromPossiblyMergedCell(currentWorksheet, i, 3);
                                        //cOO.Observation = temp.ToString(); 
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            COO.Add(cOO);

                            cOO = new CleanOrderlyOrganized();
                            for (int i = 18; i <= 29; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["C" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "C";

                                    if (i == 18)
                                    {
                                        cOO.A1 = temp.ToString();
                                    }
                                    else if (i == 20)
                                    {
                                        cOO.A2 = temp.ToString();
                                    }
                                    else if (i == 22)
                                    {
                                        cOO.A3 = temp.ToString();
                                    }
                                    else if (i == 24)
                                    {
                                        cOO.A4 = temp.ToString();
                                    }
                                    else if (i == 26)
                                    {
                                        cOO.A5 = temp.ToString();
                                    }
                                    else if (i == 29)
                                    {
                                        cOO.Observation = GetCellValueFromPossiblyMergedCell(currentWorksheet, i, 3);
                                        //cOO.Observation = temp.ToString(); 
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            COO2.Add(cOO);

                            cOO = new CleanOrderlyOrganized();
                            for (int i = 34; i <= 41; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["C" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "C";

                                    if (i == 34)
                                    {
                                        cOO.A1 = temp.ToString();
                                    }
                                    else if (i == 36)
                                    {
                                        cOO.A2 = temp.ToString();
                                    }
                                    else if (i == 38)
                                    {
                                        cOO.A3 = temp.ToString();
                                    }
                                    else if (i == 41)
                                    {
                                        cOO.Observation = GetCellValueFromPossiblyMergedCell(currentWorksheet, i, 3);
                                        //cOO.Observation = temp.ToString(); 
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            COO3.Add(cOO);

                            cOO = new CleanOrderlyOrganized();
                            for (int i = 46; i <= 57; i++)
                            {
                                try
                                {
                                    var temp = currentWorksheet.Cells["C" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "C";

                                    if (i == 46)
                                    {
                                        cOO.A1 = temp.ToString();
                                    }
                                    else if (i == 48)
                                    {
                                        cOO.A2 = temp.ToString();
                                    }
                                    else if (i == 50)
                                    {
                                        cOO.A3 = temp.ToString();
                                    }
                                    else if (i == 52)
                                    {
                                        cOO.A4 = temp.ToString();
                                    }
                                    else if (i == 54)
                                    {
                                        cOO.A5 = temp.ToString();
                                    }
                                    else if (i == 57)
                                    {
                                        cOO.Observation = GetCellValueFromPossiblyMergedCell(currentWorksheet, i, 3);
                                        //cOO.Observation = temp.ToString(); 
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            COO4.Add(cOO);

                            cOO = new CleanOrderlyOrganized();
                            for (int i = 61; i <= 71; i++)
                            {
                                try
                                {
                                    var temp = currentWorksheet.Cells["C" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "C";

                                    if (i == 62)
                                    {
                                        cOO.A1 = temp.ToString();
                                    }
                                    else if (i == 64)
                                    {
                                        cOO.A2 = temp.ToString();
                                    }
                                    else if (i == 66)
                                    {
                                        cOO.A3 = temp.ToString();
                                    }
                                    else if (i == 68)
                                    {
                                        cOO.A4 = temp.ToString();
                                    }
                                    else if (i == 71)
                                    {
                                        cOO.Observation = GetCellValueFromPossiblyMergedCell(currentWorksheet, i, 3);
                                        //cOO.Observation = temp.ToString(); 
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            COO5.Add(cOO);

                            cOO = new CleanOrderlyOrganized();
                            for (int i = 76; i <= 87; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["C" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "C";

                                    if (i == 76)
                                    {
                                        cOO.A1 = temp.ToString();
                                    }
                                    else if (i == 78)
                                    {
                                        cOO.A2 = temp.ToString();
                                    }
                                    else if (i == 80)
                                    {
                                        cOO.A3 = temp.ToString();
                                    }
                                    else if (i == 82)
                                    {
                                        cOO.A4 = temp.ToString();
                                    }
                                    else if (i == 84)
                                    {
                                        cOO.A5 = temp.ToString();
                                    }
                                    else if (i == 87)
                                    {
                                        cOO.Observation = GetCellValueFromPossiblyMergedCell(currentWorksheet, i, 3);
                                        //cOO.Observation = temp.ToString(); 
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            COO6.Add(cOO);

                            cOO = new CleanOrderlyOrganized();
                            for (int i = 92; i <= 102; i++)
                            {
                                try
                                { 
                                    var temp = currentWorksheet.Cells["C" + i + ""].Value;

                                    RowIdx = i;
                                    ColIdx = "C";

                                    if (i == 92)
                                    {
                                        cOO.A1 = temp.ToString();
                                    }
                                    else if (i == 94)
                                    {
                                        cOO.A2 = temp.ToString();
                                    }
                                    else if (i == 96)
                                    {
                                        cOO.A3 = temp.ToString();
                                    }
                                    else if (i == 98)
                                    {
                                        cOO.A4 = temp.ToString();
                                    }
                                    else if (i == 101)
                                    {
                                        cOO.A5 = temp.ToString();
                                    }
                                    else if (i == 102)
                                    {
                                        cOO.Observation = GetCellValueFromPossiblyMergedCell(currentWorksheet, i, 3);
                                        //cOO.Observation = temp.ToString(); 
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(" error happen at sheet " + currentWorksheet.Name + " Full address " + currentWorksheet.Cells.FullAddressAbsolute + "The Error its at Row = " + RowIdx + " Col = " + ColIdx + " /n The Exception is /n" + ex.ToString());
                                }
                            }
                            COO7.Add(cOO);



                            pageCount++;
                        }

                             


                        //foreach (var page in workBook.Worksheets)
                        //{
                        //    ExcelWorksheet currentWorksheet = page;

                        //for (int i = currentWorksheet.Dimension.Start.Row;
                        //    i <= currentWorksheet.Dimension.End.Row;
                        //    i++)
                        //{
                        //    for (int j = currentWorksheet.Dimension.Start.Column;
                        //             j <= currentWorksheet.Dimension.End.Column;
                        //             j++)
                        //    {
                        //        object cellValue = currentWorksheet.Cells[i, j].Value;
                        //        //ViewBag.pront = cellValue.ToString();
                        //        opj.Add(cellValue);
                        //    }
                        //}
                        //}





                        //foreach (Excel.Worksheet cws in currentWorksheet.Cells)
                        //{
                        //    var x = cws.Next;
                        //    // or whatever you want to do with the worksheet	
                        //}
                        //this is important to hold onto the range reference
                        //var cells = currentWorksheet.Cells;

                        //this is important to start the cellEnum object (the Enumerator)
                        //cells.Reset();

                        //while (cells.MoveNext())
                        //{
                        //    //Current can now be used thanks to MoveNext
                        //    Console.WriteLine("Cell [{0}, {1}] = {2}"
                        //        , cells.Current.Start.Row
                        //        , cells.Current.Start.Column
                        //        , cells.Current.Value);
                        //}


                        // gets the currentWorksheet but doesn't evaluate anything...
                        //Excel._Worksheet xlWorksheet = xlWorkbook.Worksheets.Sheets[1];
                        //Excel.Range xlRange2 = currentWorksheet.UsedRange;

                        //int rowCount2 = xlRange2.Rows.Count;
                        //int colCount2 = xlRange2.Columns.Count;


                    }
                }
            }
        }

        public string GetCellValueFromPossiblyMergedCell(ExcelWorksheet wks, int row, int col)
        {
            var cell = wks.Cells[row, col];
            if (cell.Merge)
            {
                var mergedId = wks.MergedCells[row, col];
                return wks.Cells[mergedId].First().Value.ToString();
            }
            else
            {
                return cell.Value.ToString();
            }
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelFile)
        {
            ViewBag.dep = "Work";
            if (excelFile == null || excelFile.ContentLength == 0)
            {
                ViewBag.Error = "Please select an Excel File<br>";
                return View("Index");
            }
            else if (excelFile.FileName.EndsWith("xls") || excelFile.FileName.EndsWith("xlsx"))
            {
                string path = Server.MapPath("~/Content/"+excelFile.FileName);
                if (System.IO.File.Exists(path))
                    System.IO.File.Delete(path);
                excelFile.SaveAs(path);
                ReadExcelFile();
                ViewBag.dep = "Work";
                return View("Index");

            }
            else
            {
                ViewBag.Error = "File type is incorrect<br>";
                return View("Index");
            }
        }
    }
}