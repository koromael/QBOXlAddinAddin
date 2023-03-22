using Intuit.Ipp.OAuth2PlatformClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace QBOXlAddIn.QBO
{
    class GetToken
    {

        public static string RealmId = "";
        public static string Access_token = "";
        public static string Refresh_token = "";

        public static string clientid = ConfigurationManager.AppSettings["clientid"];
        public static string clientsecret = ConfigurationManager.AppSettings["clientsecret"];
        public static string redirectUrl = ConfigurationManager.AppSettings["redirectUrl"];
        public static string environment = ConfigurationManager.AppSettings["appEnvironment"];
        public static string QboBaseUrl = ConfigurationManager.AppSettings["QboBaseUrl"];
        static OAuth2Client oauthClient = new OAuth2Client(clientID, clientSecret, redirectURI, appEnvironment);

        public ActionResult InitiateAuth()
        {
            List<OidcScopes> scopes = new List<OidcScopes>();
            scopes.Add(OidcScopes.Accounting);
            string authorizeUrl = oauthClient.GetAuthorizationURL(scopes);
            return Redirect(authorizeUrl);
        }

        public async Task<ActionResult> QboCallBack()
        {
            string code = Request.QueryString["code"] ?? "none";
            string realmId = Request.QueryString["realmId"] ?? "none";
            if (code != "none" && realmId != "none")
            {
                RealmId = realmId;
                var tokenResponse = await auth2Client.GetBearerTokenAsync(code);
                if (!string.IsNullOrWhiteSpace(tokenResponse.AccessToken))
                {
                    Access_token = tokenResponse.AccessToken;
                }
                if (!string.IsNullOrWhiteSpace(tokenResponse.RefreshToken))
                {
                    Refresh_token = tokenResponse.RefreshToken;
                }
            }
            return RedirectToAction("index", "Home");
        }


    }
}
