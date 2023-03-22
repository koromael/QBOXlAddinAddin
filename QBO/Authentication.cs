using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.OAuth2PlatformClient;
using Intuit.Ipp.ReportService;
using Intuit.Ipp.Security;
using Intuit.Ipp.Diagnostics;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;

namespace QBOXlAddIn.QBO
{
    class Authentication
    {
        static string baseURL = ConfigurationManager.AppSettings["baseURL"];

        public ReportService QboReportService(OAuth2RequestValidator authvalidator)
        {
            OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(AppGlobals.ACCESS_TOKEN);

            ServiceContext serviceContext = new ServiceContext(AppGlobals.REALM_ID, IntuitServicesType.QBO, oauthValidator);
            serviceContext.IppConfiguration.MinorVersion.Qbo = "23";
            serviceContext.IppConfiguration.BaseUrl.Qbo = baseURL;

            //JSON required for QBO Reports API
            serviceContext.IppConfiguration.Message.Response.SerializationFormat = Intuit.Ipp.Core.Configuration.SerializationFormat.Json;

            //Instantiate ReportService
            ReportService reportsService = new ReportService(serviceContext);

            return reportsService;
        }

        string getAuthorisationRequestURL()
            {
                List<OidcScopes> scopes = new List<OidcScopes>();
                scopes.Add(OidcScopes.Accounting);
                var authorizationRequest = oauthClient.GetAuthorizationURL(scopes);
                string reqdAuthorizationRequest = authorizationRequest.ToString();
                if (string.IsNullOrEmpty(reqdAuthorizationRequest))
                {
                    return "";
                }
                else
                {
                    return reqdAuthorizationRequest;
                }

            }
        
    }

}
