using Intuit.Ipp.OAuth2PlatformClient;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Net;
using System.Windows.Automation;
using System.Windows.Forms;

namespace QBOXlAddIn
{
    public partial class Ribbon1
    {
        static string redirectURI = ConfigurationManager.AppSettings["redirectURI"];
        static string clientID = ConfigurationManager.AppSettings["clientID"];
        static string clientSecret = ConfigurationManager.AppSettings["clientSecret"];
        static string logPath = ConfigurationManager.AppSettings["logPath"];
        static string appEnvironment = ConfigurationManager.AppSettings["appEnvironment"];
        static string baseURL = ConfigurationManager.AppSettings["baseURL"];

        static OAuth2Client oauthClient = new OAuth2Client(clientID, clientSecret, redirectURI, appEnvironment);

        HttpListener QBOHttpListener = null;
        Process browserProcess = null;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (QBOHttpListener == null)
                QBOHttpListener = new HttpListener();
            else
            {
                if (QBOHttpListener.IsListening)
                {
                    QBOHttpListener.Stop();
                    QBOHttpListener.Abort();
                }
                QBOHttpListener = new HttpListener();
            }
            QBOHttpListener.Prefixes.Add(redirectURI); //"http://localhost:33337/";
            QBOHttpListener.Start();

            List<OidcScopes> scopes = new List<OidcScopes>();
            scopes.Add(OidcScopes.Accounting);
            var authorizationRequest = oauthClient.GetAuthorizationURL(scopes);

            //Browser b = new Browser();
            //b.Navigate(authorizationRequest.ToString());
            //Navigate(authorizationRequest.ToString());

            browserProcess = new Process();
            browserProcess.StartInfo.FileName = authorizationRequest.ToString();
            //browserProcess.StartInfo.UseShellExecute = true;
            //browserProcess.StartInfo.Verb = "";
            //browserProcess.StartInfo.Arguments = browserProcess.StartInfo.FileName + " --new-window";
            //browserProcess.StartInfo.Arguments = "--window-size" + " " + "400,200";
            try
            {
                //Launch Chrome in a new window
                Process.Start("chrome", authorizationRequest.ToString() + " --new-window");
            }
            catch
            {
                try
                {
                    //Chrome not found ... launch Firefox in a new window
                    Process.Start("firefox", "-new-window " + authorizationRequest.ToString());
                }
                catch
                {
                    browserProcess.Start();
                }
            }




            var context = QBOHttpListener.GetContext();
            var response = context.Response;
            string responseString;
            if (response != null)
            {
                responseString = string.Format("<html><head></head><body>You are connected to quickbooks online!</body></html>");
            }
            else
            {
                responseString = string.Format("<html><head></head><body>Failed to connect to quickbooks online!!!</body></html>");
            }
            var buffer = System.Text.Encoding.UTF8.GetBytes(responseString);
            response.ContentLength64 = buffer.Length;
            var responseOutput = response.OutputStream;
            System.Threading.Tasks.Task responseTask = responseOutput.WriteAsync(buffer, 0, buffer.Length).ContinueWith((task) =>
            {
                responseOutput.Close();
                QBOHttpListener.Stop();
                QBOHttpListener = null;
            });

            CloseAllChromeBrowsers();

            AppGlobals.CODE = context.Request.QueryString.Get("code");
            AppGlobals.STATE = context.Request.QueryString.Get("state");
            AppGlobals.REALM_ID = context.Request.QueryString.Get("realmid");

            var tokenResponse = oauthClient.GetBearerTokenAsync(AppGlobals.CODE).Result;

            AppGlobals.ACCESS_TOKEN = tokenResponse.AccessToken;
            AppGlobals.REFRESH_TOKEN = tokenResponse.RefreshToken;
            long ACCESS_TOKEN_EXPIRES_IN = tokenResponse.RefreshTokenExpiresIn;

            //dynamic worksheet = AppGlobals.ThisAddIn.Application.ActiveSheet;


            Form1 Frm = new Form1();
            Frm.Show();
            Frm.BringToFront();



        }
        //public void CloseActiveBrowser()
        //{
        //    System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("CHROME");

        //    foreach (System.Diagnostics.Process proc in procs)
        //    {
        //        // Look for Google title.
        //        string s = GetBrowsedUrl(proc);
        //        string subS = redirectURI;
        //        //if (proc.MainWindowTitle.ToLower().IndexOf("localhost") > -1)
        //        if (s != null)
        //        {
        //            if (s.Contains(subS.Substring(subS.Length - 16)))
        //            {
        //                proc.Kill(); // Close it down.
        //                break;
        //            }
        //        }
        //    }
        //}

        static void CloseAllChromeBrowsers()
        {
            foreach (Process process in Process.GetProcessesByName("chrome"))
            {
                if (process.MainWindowHandle == IntPtr.Zero) // some have no UI
                    continue;

                AutomationElement element = AutomationElement.FromHandle(process.MainWindowHandle);
                if (element != null)
                {
                    string s = GetBrowserUrl(process);
                    string subS = redirectURI;
                    //if (proc.MainWindowTitle.ToLower().IndexOf("localhost") > -1)
                    if (s != null)
                    {
                        if (s.Contains(subS.Substring(subS.Length - 16)))
                        {
                            ((WindowPattern)element.GetCurrentPattern(WindowPattern.Pattern)).Close();
                            break;
                        }
                    }
                }
            }
            foreach (Process process in Process.GetProcessesByName("firefox"))
            {
                if (process.MainWindowHandle == IntPtr.Zero) // some have no UI
                    continue;

                AutomationElement element = AutomationElement.FromHandle(process.MainWindowHandle); 
                if (element != null)
                {
                    string s = GetBrowserUrl(process);
                    string subS = redirectURI;
                    if (s.Contains(subS.Substring(subS.Length - 16)))
                    {
                        ((WindowPattern)element.GetCurrentPattern(WindowPattern.Pattern)).Close();
                    }
                }
            }
        }

        public static string GetBrowserUrl(Process process)
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

    }
}


