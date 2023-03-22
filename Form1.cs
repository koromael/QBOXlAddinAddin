using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.ReportService;
using Intuit.Ipp.Security;
using System;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;


namespace QBOXlAddIn
{
    public partial class Form1 : Form
    {


        static string baseURL = ConfigurationManager.AppSettings["baseURL"];
        public Form1()
        {
            InitializeComponent();
            InitializeFormEntries();
        }

        private void InitializeFormEntries()
        {

            //this.Show();
            //this.WindowState = FormWindowState.Normal;
            //this.BringToFront();
            //this.TopLevel = true;
            //this.Focus();
            this.dateTimePicker1.CustomFormat = " ";
            this.dateTimePicker2.CustomFormat = " ";
            AppGlobals.REPORT_START_DATE = "";
            AppGlobals.REPORT_END_DATE = "";
    }


        //public void OAuthRequestInitialize()
        //{
        //    OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(AppGlobals.ACCESS_TOKEN);
        //    // Create a ServiceContext with Auth tokens and realmId
        //    ServiceContext serviceContext = new ServiceContext(AppGlobals.REALM_ID, IntuitServicesType.QBO, oauthValidator);
        //    serviceContext.IppConfiguration.MinorVersion.Qbo = "23";
        //    serviceContext.IppConfiguration.BaseUrl.Qbo = baseURL;

        //    //JSON required for QBO Reports API
        //    serviceContext.IppConfiguration.Message.Response.SerializationFormat = Intuit.Ipp.Core.Configuration.SerializationFormat.Json;

        //    //Instantiate ReportService
        //    ReportService reportsService = new ReportService(serviceContext);

        //}

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public void button1_Click(object sender, EventArgs e)
        {
            
            //var principal = System.Security.Principal.IPrincipal User as ClaimsPrincipal;
            //User = context.User;
            OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(AppGlobals.ACCESS_TOKEN);
            // Create a ServiceContext with Auth tokens and realmId
            ServiceContext serviceContext = new ServiceContext(AppGlobals.REALM_ID, IntuitServicesType.QBO, oauthValidator);
            serviceContext.IppConfiguration.MinorVersion.Qbo = "23";
            serviceContext.IppConfiguration.BaseUrl.Qbo = baseURL;

            //JSON required for QBO Reports API
            serviceContext.IppConfiguration.Message.Response.SerializationFormat = Intuit.Ipp.Core.Configuration.SerializationFormat.Json;

            //Instantiate ReportService
            ReportService reportsService = new ReportService(serviceContext);

            //Set properties for Report
            if (AppGlobals.REPORT_START_DATE != "" && AppGlobals.REPORT_END_DATE != "")
            {
                reportsService.start_date = AppGlobals.REPORT_START_DATE;
                reportsService.end_date = AppGlobals.REPORT_START_DATE;

            }
            String reportName = "ProfitAndLoss";
            //reportsService.customer = CustomerID;

            //Execute Report API call
            Report report = reportsService.ExecuteReport(reportName);
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

            DumpDataToExcelSheet(arrBSData2, 2, arrCount);

        }

        private void button2_Click(object sender, EventArgs e)
        {

            //var principal = System.Security.Principal.IPrincipal User as ClaimsPrincipal;
            //User = context.User;
            OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(AppGlobals.ACCESS_TOKEN);
            // Create a ServiceContext with Auth tokens and realmId
            ServiceContext serviceContext = new ServiceContext(AppGlobals.REALM_ID, IntuitServicesType.QBO, oauthValidator);
            serviceContext.IppConfiguration.MinorVersion.Qbo = "23";
            serviceContext.IppConfiguration.BaseUrl.Qbo = baseURL;

            //JSON required for QBO Reports API
            serviceContext.IppConfiguration.Message.Response.SerializationFormat = Intuit.Ipp.Core.Configuration.SerializationFormat.Json;

            //Instantiate ReportService
            ReportService reportsService = new ReportService(serviceContext);

            //Set properties for Report
            if (AppGlobals.REPORT_START_DATE != "" && AppGlobals.REPORT_END_DATE != "")
            {
                reportsService.start_date = AppGlobals.REPORT_START_DATE;
                reportsService.end_date = AppGlobals.REPORT_START_DATE;

            }
            String reportName = "BalanceSheet";
            //reportsService.customer = CustomerID;

            //Execute Report API call
            Report report = reportsService.ExecuteReport(reportName);

            string ReportStr = string.Empty;
            XmlConverter str = new XmlConverter();
            ReportStr = str.getReportInXML(report);

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(ReportStr);

            XmlNodeList xmlnode;
            xmlnode = xmlDoc.GetElementsByTagName("Rows");

            string[,] arrBSData = new string[10000, 2];
            arrBSData[0, 0] = "BALANBCESHEET REPORT";
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

            DumpDataToExcelSheet(arrBSData2, 2, arrCount);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(AppGlobals.ACCESS_TOKEN);

            ServiceContext serviceContext = new ServiceContext(AppGlobals.REALM_ID, IntuitServicesType.QBO, oauthValidator);
            serviceContext.IppConfiguration.MinorVersion.Qbo = "23";
            serviceContext.IppConfiguration.BaseUrl.Qbo = baseURL;
            serviceContext.IppConfiguration.Message.Response.SerializationFormat = Intuit.Ipp.Core.Configuration.SerializationFormat.Json;
            ReportService reportsService = new ReportService(serviceContext);

            //Set properties for Report
            if (AppGlobals.REPORT_START_DATE != "" && AppGlobals.REPORT_END_DATE != "")
            {
                reportsService.start_date = AppGlobals.REPORT_START_DATE;
                reportsService.end_date = AppGlobals.REPORT_START_DATE;

            }
            String reportName = "TrialBalance";
            //reportsService.customer = CustomerID;

            //Execute Report API call
            Report report = reportsService.ExecuteReport(reportName);

            string ReportStr = string.Empty;
            XmlConverter str = new XmlConverter();
            ReportStr = str.getReportInXML(report);

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(ReportStr);
            XmlNodeList xmlnode;
            xmlnode = xmlDoc.GetElementsByTagName("Rows");

            string[,] arrTBData = new string[10000, 3];
            arrTBData[0, 0] = "TRIAL BALANBCESHEET REPORT";
            arrTBData[0, 1] = "";
            arrTBData[0, 2] = "";
            arrTBData[1, 0] = "";
            arrTBData[1, 1] = "";
            arrTBData[1, 2] = "";
            arrTBData[2, 0] = "ACCOUNT";
            arrTBData[2, 1] = "DEBIT";
            arrTBData[2, 2] = "CREDIT"; 
            int colcount = 3;
            int rowcount = report.Rows.Count() + 3;

            for (int i = 3; i < rowcount; i++)
            {
                if (i == rowcount - 1)
                {
                    arrTBData[i, 0] = ((Summary)report.Rows[i - 3].AnyIntuitObjects[0]).ColData[0].value;
                    arrTBData[i, 1] = ((Summary)report.Rows[i - 3].AnyIntuitObjects[0]).ColData[1].value;
                    arrTBData[i, 2] = ((Summary)report.Rows[i - 3].AnyIntuitObjects[0]).ColData[2].value;
                }
                else
                {
                    arrTBData[i, 0] = ((Intuit.Ipp.Data.ColData[])report.Rows[i - 3].AnyIntuitObjects[0])[0].value;
                    arrTBData[i, 1] = ((Intuit.Ipp.Data.ColData[])report.Rows[i - 3].AnyIntuitObjects[0])[1].value;
                    arrTBData[i, 2] = ((Intuit.Ipp.Data.ColData[])report.Rows[i - 3].AnyIntuitObjects[0])[2].value;
                }
            }
            DumpDataToExcelSheet(arrTBData, colcount, rowcount);

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void DumpDataToExcelSheet(string[,] DataToDump, int ColCount, int RowCount)
        {
            //object activeSheet = Globals.ThisAddIn.Application.ThisWorkbook.ActiveSheet;

            //dynamic cell = Globals.ThisAddIn.Application.ActiveCell;

            Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell;
            int row = rng.Row - 1;
            int column = rng.Column - 1;

            const String CELL = "A1";

            for (int i = 0; i <= RowCount; i++)
            {
                for (int j = 1; j <= ColCount; j++)
                {
                    if (DataToDump[i, 0] == null && DataToDump[i , 1] == null)
                    {
                        break;
                    }
                    Globals.ThisAddIn.Application.ActiveSheet.Range(CELL).Offset(row + i, column + j).Value = DataToDump[i , j - 1];
                }
            }
            Globals.ThisAddIn.Application.ActiveSheet.Columns.AutoFit();
        }

        public string GetBrowsedUrl(Process process)
        {
            if (process.ProcessName == "firefox")
            {
                if (process == null)
                    throw new ArgumentNullException("process");

                if (process.MainWindowHandle == IntPtr.Zero)
                    return null;

                AutomationElement element = AutomationElement.FromHandle(process.MainWindowHandle);
                if (element == null)
                    return null;

                AutomationElement doc = element.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document));
                if (doc == null)
                    return null;

                return ((ValuePattern)doc.GetCurrentPattern(ValuePattern.Pattern)).Current.Value as string;
            }
            else
            {
                if (process == null)
                    throw new ArgumentNullException("process");

                if (process.MainWindowHandle == IntPtr.Zero)
                    return null;

                AutomationElement element = AutomationElement.FromHandle(process.MainWindowHandle);
                if (element == null)
                    return null;

                AutomationElement edit = element.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                string result = ((ValuePattern)edit.GetCurrentPattern(ValuePattern.Pattern)).Current.Value as string;
                return result;
            }

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            AppGlobals.REPORT_START_DATE = this.dateTimePicker1.Value.ToString("YYYYMMDD"); 
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            AppGlobals.REPORT_END_DATE = this.dateTimePicker1.Value.ToString("YYYYMMDD"); 
        }
    }
}
