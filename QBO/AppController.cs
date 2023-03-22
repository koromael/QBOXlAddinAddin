using Intuit.Ipp.OAuth2PlatformClient;
using System;
using System.Security.Claims;
using System.Web;
using System.Web.Mvc;
using System.Net;
using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.QueryFilter;
using Intuit.Ipp.Security;
using System.Linq;
using System.Collections.Generic;
using Intuit.Ipp.ReportService;
using System.Text;
using MvcCodeFlowClientManual.Helper;
using System.Xml;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Header = Intuit.Ipp.Data.Header;
using Row = Intuit.Ipp.Data.Row;
using Column = Intuit.Ipp.Data.Column;
using ConfigurationManager = System.Configuration.ConfigurationManager;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using System.Diagnostics;

namespace MvcCodeFlowClientManual.Controllers
{
    public class AppController : Controller
    {
        public static string clientid = ConfigurationManager.AppSettings["clientid"];
        public static string clientsecret = ConfigurationManager.AppSettings["clientsecret"];
        public static string redirectUrl = ConfigurationManager.AppSettings["redirectUrl"];
        public static string environment = ConfigurationManager.AppSettings["appEnvironment"];

        public static OAuth2Client auth2Client = new OAuth2Client(clientid, clientsecret, redirectUrl, environment);

        /// <summary>
        /// Use the Index page of App controller to get all endpoints from discovery url
        /// </summary>
        public ActionResult Index()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            Session.Clear();
            Session.Abandon();
            Request.GetOwinContext().Authentication.SignOut("Cookies");
            return View();
        }

        /// <summary>
        /// Start Auth flow
        /// </summary>
        public ActionResult InitiateAuth(string submitButton)
        {
            switch (submitButton)
            {
                case "Connect to QuickBooks":
                    List<OidcScopes> scopes = new List<OidcScopes>();
                    scopes.Add(OidcScopes.Accounting);
                    string authorizeUrl = auth2Client.GetAuthorizationURL(scopes);
                    return Redirect(authorizeUrl);
                default:
                    return (View());
            }
        }

        /// <summary>
        /// QBO API Request
        /// </summary>
        public ActionResult ApiCallService()
        {
            if (Session["realmId"] != null)
            {
                string realmId = Session["realmId"].ToString();
                try
                {
                    var principal = User as ClaimsPrincipal;
                    OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(principal.FindFirst("access_token").Value);

                    // Create a ServiceContext with Auth tokens and realmId
                    ServiceContext serviceContext = new ServiceContext(realmId, IntuitServicesType.QBO, oauthValidator);
                    serviceContext.IppConfiguration.MinorVersion.Qbo = "23";

                    // Create a QuickBooks QueryService using ServiceContext
                    QueryService<CompanyInfo> querySvc = new QueryService<CompanyInfo>(serviceContext);
                    CompanyInfo companyInfo = querySvc.ExecuteIdsQuery("SELECT * FROM CompanyInfo").FirstOrDefault();

                    string output = "Company Name: " + companyInfo.CompanyName + " Company Address: " + companyInfo.CompanyAddr.Line1 + ", " + companyInfo.CompanyAddr.City + ", " + companyInfo.CompanyAddr.Country + " " + companyInfo.CompanyAddr.PostalCode;
                    return View("ApiCallService", (object)("QBO API call Successful!! Response: " + output));
                }
                catch (Exception ex)
                {
                    return View("ApiCallService", (object)("QBO API call Failed!" + " Error message: " + ex.Message));
                }
            }
            else
                return View("ApiCallService", (object)"QBO API call Failed!");
        }

        /// <summary>
        /// QBO API Request
        /// </summary>
        public ActionResult PnLCallService() //GetCustomerProfitLossReport()
        {
            if (Session["realmId"] != null)
            {
                string realmId = Session["realmId"].ToString();
                ProfitLossReport CustPLReport = new ProfitLossReport();
                try
                {
                    var principal = User as ClaimsPrincipal;
                    OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(principal.FindFirst("access_token").Value);
                    // Create a ServiceContext with Auth tokens and realmId
                    ServiceContext serviceContext = new ServiceContext(realmId, IntuitServicesType.QBO, oauthValidator);
                    serviceContext.IppConfiguration.MinorVersion.Qbo = "23";
                    //serviceContext.IppConfiguration.BaseUrl.Qbo = QboBaseUrl;


                    //JSON required for QBO Reports API
                    serviceContext.IppConfiguration.Message.Response.SerializationFormat = Intuit.Ipp.Core.Configuration.SerializationFormat.Json;

                    //Instantiate ReportService
                    ReportService reportsService = new ReportService(serviceContext);

                    //Set properties for Report
                    //reportsService.start_date = startDate;
                    //reportsService.end_date = endDate;
                    String reportName = "ProfitAndLoss";
                    //reportsService.customer = CustomerID;


                    //Execute Report API call
                    Report report = reportsService.ExecuteReport(reportName);
                    string ReportStr = string.Empty;

                    //Format the report data 
                    // ReportStr = PrintReportToString(report);

                    // AddNamedRange();
                    string ReportStr2 = string.Empty;
                    XmlConverter str = new XmlConverter();
                    ReportStr2 = str.getReportInXML(report);

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(ReportStr2);

                    XmlNodeList xmlnode;
                    xmlnode = xmlDoc.GetElementsByTagName("Rows");

                    string[,] arrPnLData = new string[10000, 2];
                    arrPnLData[0, 0] = "PROFIT AND LOSS REPORT";
                    arrPnLData[0, 1] = "";
                    arrPnLData[1, 0] = "";
                    arrPnLData[1, 1] = "";
                    int arrCount = 2;


                    int i = 0;
                    for (int j = 0; j < xmlnode[i].ChildNodes.Count; j++)
                    {
                        arrPnLData[arrCount, 0] = xmlnode[i].ChildNodes[j].Attributes[1].InnerText;
                        arrPnLData[arrCount, 1] = "";
                        arrCount++;

                        for (int k = 0; k < xmlnode[i].ChildNodes[j].ChildNodes.Count; k++)
                        {
                            if (xmlnode[i].ChildNodes[j].ChildNodes[k].Name == "Header")
                            {
                                arrPnLData[arrCount - 1, 1] = xmlnode[i].ChildNodes[j].ChildNodes[k].ChildNodes[0].Attributes[0].Value;
                            }
                            else if (xmlnode[i].ChildNodes[j].ChildNodes[k].Name == "Rows")
                            {
                                for (int l = 0; l < xmlnode[i].ChildNodes[j].ChildNodes[k].ChildNodes.Count; l++)
                                {
                                    arrPnLData[arrCount, 0] = xmlnode[i].ChildNodes[j].ChildNodes[k].ChildNodes[l].ChildNodes[0].Attributes[0].Value;
                                    arrPnLData[arrCount, 1] = xmlnode[i].ChildNodes[j].ChildNodes[k].ChildNodes[l].ChildNodes[1].Attributes[0].Value;
                                    arrCount++;
                                }
                            }
                            else if (xmlnode[i].ChildNodes[j].ChildNodes[k].Name == "Summary")
                            {
                                arrPnLData[arrCount, 0] = xmlnode[i].ChildNodes[j].ChildNodes[k].ChildNodes[0].Attributes[0].Value;
                                arrPnLData[arrCount, 1] = xmlnode[i].ChildNodes[j].ChildNodes[k].ChildNodes[1].Attributes[0].Value;
                                arrCount++;
                            }
                        }
                    }
                    DumpDataToExcelSheet(arrPnLData, 2, 1000);

                    //   CustPLReport.ID = "A";
                    //   CustPLReport.Name = "PnL";
                    //   CustPLReport.ReportText = ReportStr;

                }
                catch (Exception ex)
                {

                }
                return View("ApiCallService", (object)CustPLReport.ReportText);
            }
            else
                return View("ApiCallService", (object)"QBO API call Failed!");
        }


        private void DumpDataToExcelSheet(string[,] dumpData, int ColCount, int RowCount)
        {
            const string EXCEL_PROG_ID = "Excel.Application";

            const uint MK_E_UNAVAILABLE = 0x800401e3;

            const uint DV_E_FORMATETC = 0x80040064;

            dynamic excelApp = null;
            try
            {
                excelApp = Marshal.GetActiveObject(EXCEL_PROG_ID);
            }
            catch (COMException ex)
            {
                switch ((uint)ex.ErrorCode)
                {
                    case MK_E_UNAVAILABLE:
                    case DV_E_FORMATETC:
                        // Excel n'est pas lancé.
                        break;

                    default:
                        throw;
                }
            }

            if (null == excelApp)
                excelApp = Activator.CreateInstance(Type.GetTypeFromProgID(EXCEL_PROG_ID));

            if (null == excelApp)
            {
                Console.Write("Unable to start Excel");
                return;
            }

            if (!excelApp.Visible)
            {
                excelApp.Visible = true;
            }
            dynamic workbook = excelApp.ActiveWorkbook ?? excelApp.Workbooks.Add();
            dynamic sheet = workbook.ActiveSheet;
           // dynamic cell = sheet.Cells[1, 1];

            for (int i = 1; i <= RowCount; i++)
            {
                for (int j = 1; j <= ColCount; j++)
                {
                    if (dumpData[i - 1, 0] == null && dumpData[i - 1, 1] == null)
                    {
                         break;
                    }
                    sheet.Cells[i, j].Value = dumpData[i - 1, j - 1];
                }
            }
            sheet.Columns.AutoFit();
   

        }



        /// <summary>
        /// QBO API Request
        /// </summary>
        public ActionResult BSCallService() //GetCustomerBalanceSheetReport()
        {
            if (Session["realmId"] != null)
            {
                string realmId = Session["realmId"].ToString();
                BalancesheetReport CustBSReport = new BalancesheetReport();
                try
                {
                    var principal = User as ClaimsPrincipal;
                    OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(principal.FindFirst("access_token").Value);
                    // Create a ServiceContext with Auth tokens and realmId
                    ServiceContext serviceContext = new ServiceContext(realmId, IntuitServicesType.QBO, oauthValidator);
                    serviceContext.IppConfiguration.MinorVersion.Qbo = "23";
                    //serviceContext.IppConfiguration.BaseUrl.Qbo = QboBaseUrl;


                    //JSON required for QBO Reports API
                    serviceContext.IppConfiguration.Message.Response.SerializationFormat = Intuit.Ipp.Core.Configuration.SerializationFormat.Json;

                    //Instantiate ReportService
                    ReportService reportsService = new ReportService(serviceContext);

                    //Set properties for Report
                    String reportName = "BalanceSheet";
                    //reportsService.customer = CustomerID;

                    //Execute Report API call
                    Report report = reportsService.ExecuteReport(reportName);

                    string ReportStr = string.Empty;
                    XmlConverter str = new XmlConverter();
                    ReportStr = str.getReportInXML(report);

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(ReportStr);

                    CustBSReport.ID = "A";
                    CustBSReport.Name = "BalanceSheet";
                    // CustBSReport.ReportText = ReportStr;

                    XmlNodeList xmlnode;
                    xmlnode = xmlDoc.GetElementsByTagName("Rows");

                    string[,] arrBSData = new string[10000, 2];
                    arrBSData[0, 0] = "BALANBCE SHEET REPORT";
                    arrBSData[0, 1] = "";
                    arrBSData[1, 0] = "";
                    arrBSData[1, 1] = "";
                    int arrCount = 2;
 
                    XmlNode xmlnode2 = xmlnode[0];

                    string[,] arrBSData2 = new string[10000, 2];
                    string[,] arrBSSumryData = new string[10000, 2];
                    //for (int i = 0; i < xmlnode2.ChildNodes.Count; i++)
                    //{

                    XmlNodeList xmlnode3 = xmlnode[0].ChildNodes;


                    int a = 0;
                    foreach (XmlNode node in xmlnode3) // for each <testcase> node
                    {

                        foreach (XmlNode row in node.ChildNodes)
                        {
                            if (row.Name == "Rows" || row.Name == "Row")
                            {
                                foreach (XmlNode row2 in row.ChildNodes)
                                {
                                    if (row2.Name == "Rows" || row2.Name == "Row")
                                    {
                                        foreach (XmlNode row3 in row2.ChildNodes)
                                        {
                                            if (row3.Name == "Rows" || row3.Name == "Row")
                                            {
                                                foreach (XmlNode row4 in row3.ChildNodes)
                                                {
                                                    if (row4.Name == "Rows" || row4.Name == "Row")
                                                    {
                                                        foreach (XmlNode row5 in row4.ChildNodes)
                                                        {
                                                            if (row5.Name == "Rows" || row5.Name == "Row")
                                                            {
                                                                foreach (XmlNode row6 in row5.ChildNodes)
                                                                {
                                                                    if (row6.Name == "Rows" || row6.Name == "Row")
                                                                    {
                                                                        foreach (XmlNode row7 in row6.ChildNodes)
                                                                        {
                                                                            if (row7.Name == "Rows" || row7.Name == "Row")
                                                                            {
                                                                                foreach (XmlNode row8 in row7.ChildNodes)
                                                                                {
                                                                                    if (row8.Name == "Rows" || row8.Name == "Row")
                                                                                    {
                                                                                        foreach (XmlNode row9 in row8.ChildNodes)
                                                                                        {
                                                                                            if (row9.Name == "Rows" || row9.Name == "Row")
                                                                                            {
                                                                                                foreach (XmlNode row10 in row9.ChildNodes)
                                                                                                {
                                                                                                    if (row10.Name == "Rows" || row10.Name == "Row")
                                                                                                    {
                                                                                                        if (row10.HasChildNodes)
                                                                                                        {
                                                                                                            arrBSData2[arrCount, 0] = row10.ChildNodes[0].Attributes[0].Value;
                                                                                                            arrBSData2[arrCount, 1] = row10.ChildNodes[1].Attributes[0].Value;
                                                                                                            arrCount++;
                                                                                                        }
                                                                                                        else
                                                                                                        {
                                                                                                            arrBSData2[arrCount, a] = row10.Attributes[0].Value;
                                                                                                            if (a == 0)
                                                                                                            {
                                                                                                                a++;
                                                                                                            }
                                                                                                            else
                                                                                                            {
                                                                                                                a--;
                                                                                                                arrCount++;
                                                                                                            }
                                                                                                        }
                                                                                                    }
                                                                                                }
                                                                                            }
                                                                                            else
                                                                                            {
                                                                                                if (row9.HasChildNodes)
                                                                                                {
                                                                                                    arrBSData2[arrCount, 0] = row9.ChildNodes[0].Attributes[0].Value;
                                                                                                    arrBSData2[arrCount, 1] = row9.ChildNodes[1].Attributes[0].Value;
                                                                                                    arrCount++;
                                                                                                }
                                                                                                else
                                                                                                {
                                                                                                    arrBSData2[arrCount, a] = row9.Attributes[0].Value;
                                                                                                    if (a == 0)
                                                                                                    {
                                                                                                        a++;
                                                                                                    }
                                                                                                    else
                                                                                                    {
                                                                                                        a--;
                                                                                                        arrCount++;
                                                                                                    }
                                                                                                }
                                                                                            }
                                                                                        }
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        if (row8.HasChildNodes)
                                                                                        {
                                                                                            arrBSData2[arrCount, 0] = row8.ChildNodes[0].Attributes[0].Value;
                                                                                            arrBSData2[arrCount, 1] = row8.ChildNodes[1].Attributes[0].Value;
                                                                                            arrCount++;
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            arrBSData2[arrCount, a] = row8.Attributes[0].Value;
                                                                                            if (a == 0)
                                                                                            {
                                                                                                a++;
                                                                                            }
                                                                                            else
                                                                                            {
                                                                                                a--;
                                                                                                arrCount++;
                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                if (row7.HasChildNodes)
                                                                                {
                                                                                    arrBSData2[arrCount, 0] = row7.ChildNodes[0].Attributes[0].Value;
                                                                                    arrBSData2[arrCount, 1] = row7.ChildNodes[1].Attributes[0].Value;
                                                                                    arrCount++;
                                                                                }
                                                                                else
                                                                                {
                                                                                    arrBSData2[arrCount, a] = row7.Attributes[0].Value;
                                                                                    if (a == 0)
                                                                                    {
                                                                                        a++;
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        a--;
                                                                                        arrCount++;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        if (row6.HasChildNodes)
                                                                        {
                                                                            arrBSData2[arrCount, 0] = row6.ChildNodes[0].Attributes[0].Value;
                                                                            arrBSData2[arrCount, 1] = row6.ChildNodes[1].Attributes[0].Value;
                                                                            arrCount++;
                                                                        }
                                                                        else
                                                                        {
                                                                            arrBSData2[arrCount, a] = row6.Attributes[0].Value;
                                                                            if (a == 0)
                                                                            {
                                                                                a++;
                                                                            }
                                                                            else
                                                                            {
                                                                                a--;
                                                                                arrCount++;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (row5.HasChildNodes)
                                                                {
                                                                    arrBSData2[arrCount, 0] = row5.ChildNodes[0].Attributes[0].Value;
                                                                    arrBSData2[arrCount, 1] = row5.ChildNodes[1].Attributes[0].Value;
                                                                    arrCount++;
                                                                }
                                                                else
                                                                {
                                                                    arrBSData2[arrCount, a] = row5.Attributes[0].Value;
                                                                    if (a == 0)
                                                                    {
                                                                        a++;
                                                                    }
                                                                    else
                                                                    {
                                                                        a--;
                                                                        arrCount++;
                                                                    }
                                                                            
                                                                }
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (row4.HasChildNodes)
                                                        {
                                                            arrBSData2[arrCount, 0] = row4.ChildNodes[0].Attributes[0].Value;
                                                            arrBSData2[arrCount, 1] = row4.ChildNodes[1].Attributes[0].Value;
                                                            arrCount++;
                                                        }
                                                        else
                                                        {   
                                                            arrBSData2[arrCount, a] = row4.Attributes[0].Value;
                                                            if (a == 0)
                                                            {
                                                                a++;
                                                            }
                                                            else
                                                            {
                                                                a--;
                                                                arrCount++;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (row3.HasChildNodes)
                                                {
                                                    arrBSData2[arrCount, 0] = row3.ChildNodes[0].Attributes[0].Value;
                                                    arrBSData2[arrCount, 1] = row3.ChildNodes[1].Attributes[0].Value;
                                                    arrCount++;
                                                }
                                                else
                                                {
                                                    arrBSData2[arrCount, a] = row3.Attributes[0].Value;
                                                    if (a == 0)
                                                    {
                                                        a++;
                                                    }
                                                    else
                                                    {
                                                        a--;
                                                        arrCount++;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (row2.HasChildNodes)
                                        {
                                            arrBSData2[arrCount, 0] = row2.ChildNodes[0].Attributes[0].Value;
                                            arrBSData2[arrCount, 1] = row2.ChildNodes[1].Attributes[0].Value;
                                            arrCount++;
                                        }
                                        else
                                        {
                                            arrBSData2[arrCount, a] = row2.Attributes[0].Value;
                                            if (a == 0)
                                            {
                                                a++;
                                            }
                                            else
                                            {
                                                a--;
                                                arrCount++;
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (row.HasChildNodes)
                                {
                                    arrBSData2[arrCount, 0] = row.ChildNodes[0].Attributes[0].Value;
                                    arrBSData2[arrCount, 1] = row.ChildNodes[1].Attributes[0].Value;
                                    arrCount++;
                                }
                                else
                                {
                                    arrBSData2[arrCount, a] = row.Attributes[0].Value;
                                    if (a == 0)
                                    {
                                        a++;
                                    }
                                    else
                                    {
                                        a--;
                                        arrCount++;
                                    }
                                }
                            }
                        }
                    }

                    DumpDataToExcelSheet(arrBSData2, 2, 100);
                }
                catch (Exception ex)
                {

                }
                return View("ApiCallService", (object)CustBSReport.ReportText);
            }
            else
                return View("ApiCallService", (object)"QBO API call Failed!");
        }



            /// <summary>
            /// QBO API Request
            /// </summary>
            public ActionResult TBCallService() //GetTrialBalanceSheetReport()
        {
            if (Session["realmId"] != null)
            {
                string realmId = Session["realmId"].ToString();
                TrialbalanceReport CustTBReport = new TrialbalanceReport();
                try
                {
                    var principal = User as ClaimsPrincipal;
                    OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(principal.FindFirst("access_token").Value);
                    // Create a ServiceContext with Auth tokens and realmId
                    ServiceContext serviceContext = new ServiceContext(realmId, IntuitServicesType.QBO, oauthValidator);
                    serviceContext.IppConfiguration.MinorVersion.Qbo = "23";
                    //serviceContext.IppConfiguration.BaseUrl.Qbo = QboBaseUrl;


                    //JSON required for QBO Reports API
                    serviceContext.IppConfiguration.Message.Response.SerializationFormat = Intuit.Ipp.Core.Configuration.SerializationFormat.Json;

                    //Instantiate ReportService
                    ReportService reportsService = new ReportService(serviceContext);

                    //Set properties for Report
                    //reportsService.start_date = startDate;
                    //reportsService.end_date = endDate;
                    String reportName = "TrialBalance";
                    //reportsService.customer = CustomerID;

                    //Execute Report API call
                    Report report = reportsService.ExecuteReport(reportName);
                    int colcount = 3;
                    int rowcount = report.Rows.Count();
                    string[,] arrTBData = new string[10000, 3];
                    for (int i = 0; i < rowcount; i++)
                    {
                        if (i == rowcount - 1)
                        {
                            arrTBData[i, 0] = ((Intuit.Ipp.Data.Summary)report.Rows[i].AnyIntuitObjects[0]).ColData[0].value;
                            arrTBData[i, 1] = ((Intuit.Ipp.Data.Summary)report.Rows[i].AnyIntuitObjects[0]).ColData[1].value;
                            arrTBData[i, 2] = ((Intuit.Ipp.Data.Summary)report.Rows[i].AnyIntuitObjects[0]).ColData[2].value;
                        }
                        else
                        {
                            arrTBData[i, 0] = ((Intuit.Ipp.Data.ColData[])report.Rows[i].AnyIntuitObjects[0])[0].value;
                            arrTBData[i, 1] = ((Intuit.Ipp.Data.ColData[])report.Rows[i].AnyIntuitObjects[0])[1].value;
                            arrTBData[i, 2] = ((Intuit.Ipp.Data.ColData[])report.Rows[i].AnyIntuitObjects[0])[2].value;
                        }
                    }

                    DumpDataToExcelSheet(arrTBData, colcount, rowcount);

                }
                catch (Exception ex)
                {

                }
                return View("ApiCallService", (object)CustTBReport.ReportText);
            }
            else
                return View("ApiCallService", (object)"QBO API call Failed!");
        }


        /// <summary>
        /// Use the Index page of App controller to get all endpoints from discovery url
        /// </summary>
        public ActionResult Error()
        {
            return View("Error");
        }

        /// <summary>
        /// Action that takes redirection from Callback URL
        /// </summary>
        public ActionResult Tokens()
        {
            return View("Tokens");
        }



        //private static void PrintReportToConsole(Report report)
        private string PrintReportToString(Report report)
        {
            String ReturnStr = string.Empty;
            try
            {
                StringBuilder reportText = new StringBuilder();
                //Append Report Header
                PrintHeader(reportText, report);
                //Determine Maxmimum Text Lengths to format Report
                int[] maximumColumnTextSize = GetMaximumColumnTextSize(report);
                //Append Column Headers
                PrintColumnData(reportText, report.Columns, maximumColumnTextSize, 0);
                //Append Rows
                PrintRows(reportText, report.Rows, maximumColumnTextSize, 1);
                //Formatted Report Text to Return String
                ReturnStr = reportText.ToString();
            }
            catch (Exception ex)
            {
            }
            return ReturnStr;
        }



        #region " Helper Methods "
        #region " Determine Maximum Column Text Length "
        private static int[] GetMaximumColumnTextSize(Report report)
        {
            if (report.Columns == null) { return null; }
            int[] maximumColumnSize = new int[report.Columns.Count()];
            for (int columnIndex = 0; columnIndex < report.Columns.Count(); columnIndex++)
            {
                maximumColumnSize[columnIndex] = Math.Max(maximumColumnSize[columnIndex], report.Columns[columnIndex].ColTitle.Length);
            }
            return GetMaximumRowColumnTextSize(report.Rows, maximumColumnSize, 1);
        }
        private static int[] GetMaximumRowColumnTextSize(Row[] rows, int[] maximumColumnSize, int level)
        {
            for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++)
            {
                Row row = rows[rowIndex];
                Header rowHeader = GetRowProperty<Header>(row, ItemsChoiceType1.Header);
                if (rowHeader != null) { GetMaximumColDataTextSize(rowHeader.ColData, maximumColumnSize, level); }
                ColData[] colData = GetRowProperty<ColData[]>(row, ItemsChoiceType1.ColData);
                if (colData != null) { GetMaximumColDataTextSize(colData, maximumColumnSize, level); }
                Rows nestedRows = GetRowProperty<Rows>(row, ItemsChoiceType1.Rows);
                if (nestedRows != null) { GetMaximumRowColumnTextSize(nestedRows.Row, maximumColumnSize, level + 1); }
                Summary rowSummary = GetRowProperty<Summary>(row, ItemsChoiceType1.Summary);
                if (rowSummary != null) { GetMaximumColDataTextSize(rowSummary.ColData, maximumColumnSize, level); }
            }
            return maximumColumnSize;
        }
        private static int[] GetMaximumColDataTextSize(ColData[] colData, int[] maximumColumnSize, int level)
        {
            for (int colDataIndex = 0; colDataIndex < colData.Length; colDataIndex++)
            {
                maximumColumnSize[colDataIndex] = Math.Max(maximumColumnSize[colDataIndex], (new String(' ', level * 3) + colData[colDataIndex].value).Length);
            }
            return maximumColumnSize;
        }
        #endregion
        #region " Print Report Sections "
        private static void PrintHeader(StringBuilder reportText, Report report)
        {
            const string lineDelimiter = "---";
            reportText.AppendLine(report.Header.ReportName);
            reportText.AppendLine(lineDelimiter);
            reportText.AppendLine("As of " + report.Header.StartPeriod);
            reportText.AppendLine(lineDelimiter);
            reportText.AppendLine(lineDelimiter);
        }
        private static void PrintRows(StringBuilder reportText, Row[] rows, int[] maxColumnSize, int level)
        {
            for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++)
            {
                Intuit.Ipp.Data.Row row = rows[rowIndex];
                //Get Row Header
                Header rowHeader = GetRowProperty<Header>(row, ItemsChoiceType1.Header);
                //Append Row Header
                if (rowHeader != null && rowHeader.ColData != null) { PrintColData(reportText, rowHeader.ColData, maxColumnSize, level); }
                //Get Row ColData
                ColData[] colData = GetRowProperty<ColData[]>(row, ItemsChoiceType1.ColData);
                //Append ColData
                if (colData != null) { PrintColData(reportText, colData, maxColumnSize, level); }
                //Get Child Rows
                Rows childRows = GetRowProperty<Rows>(row, ItemsChoiceType1.Rows);
                //Append Child Rows
                if (childRows != null) { PrintRows(reportText, childRows.Row, maxColumnSize, level + 1); }
                //Get Row Summary
                Summary rowSummary = GetRowProperty<Summary>(row, ItemsChoiceType1.Summary);
                //Append Row Summary
                if (rowSummary != null && rowSummary.ColData != null) { PrintColData(reportText, rowSummary.ColData, maxColumnSize, level); }
            }
        }
        private static void PrintColData(StringBuilder reportText, ColData[] colData, int[] maxColumnSize, int level)
        {
            for (int colDataIndex = 0; colDataIndex < colData.Length; colDataIndex++)
            {
                if (colDataIndex > 0) { reportText.Append("     "); }
                StringBuilder rowText = new StringBuilder();
                if (colDataIndex == 0) { rowText.Append(new String(' ', level * 3)); };
                rowText.Append(colData[colDataIndex].value);
                if (rowText.Length < maxColumnSize[colDataIndex])
                {
                    rowText.Append(new String(' ', maxColumnSize[colDataIndex] - rowText.Length));
                }
                reportText.Append(rowText.ToString());
            }
            reportText.AppendLine();
        }
        private static void PrintColumnData(StringBuilder reportText, Column[] columns, int[] maxColumnSize, int level)
        {
            for (int colDataIndex = 0; colDataIndex < columns.Length; colDataIndex++)
            {
                if (colDataIndex > 0) { reportText.Append("     "); }
                StringBuilder rowText = new StringBuilder();
                if (colDataIndex == 0) { rowText.Append(new String(' ', level * 3)); };
                rowText.Append(columns[colDataIndex].ColTitle);
                if (rowText.Length < maxColumnSize[colDataIndex])
                {
                    rowText.Append(new String(' ', maxColumnSize[colDataIndex] - rowText.Length));
                }
                reportText.Append(rowText.ToString());
            }
            reportText.AppendLine();
        }
        #endregion
        #region " Get Row Property Helper Methods - Header, ColData, Rows (children), Summary "
        //Returns typed object from AnyIntuitObjects array
        private static T GetRowProperty<T>(Intuit.Ipp.Data.Row row, ItemsChoiceType1 itemsChoiceType)
        {
            int choiceElementIndex = GetChoiceElementIndex(row, itemsChoiceType);
            if (choiceElementIndex == -1) { return default(T); } else { return (T)row.AnyIntuitObjects[choiceElementIndex]; }
        }
        //Finds element index in ItemsChoiceType array
        private static int GetChoiceElementIndex(Intuit.Ipp.Data.Row row, ItemsChoiceType1 itemsChoiceType)
        {
            if (row.ItemsElementName != null)
            {
                for (int itemsChoiceTypeIndex = 0; itemsChoiceTypeIndex < row.ItemsElementName.Count(); itemsChoiceTypeIndex++)
                {
                    if (row.ItemsElementName[itemsChoiceTypeIndex] == itemsChoiceType) { return itemsChoiceTypeIndex; }
                }
            }
            return -1;
        }
        #endregion
        #endregion









    }
}