using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace TestApp
{
    public class Token
    {
        public Token()
        {
            Issued = DateTime.Now;
        }

        [JsonProperty("access_token")]
        public string AccessToken { get; set; }

        [JsonProperty("token_type")]
        public string TokenType { get; set; }

        [JsonProperty("expires_in")]
        public int ExpiresIn { get; set; }

        [JsonProperty("refresh_token")]
        public string RefreshToken { get; set; }

        [JsonProperty("as:client_id")]
        public string ClientId { get; set; }

        [JsonProperty("userName")]
        public string UserName { get; set; }

        [JsonProperty("as:region")]
        public string Region { get; set; }

        [JsonProperty(".issued")]
        public DateTime Issued { get; set; }

        [JsonProperty(".expires")]
        public DateTime Expires
        {
            get { return Issued.AddMilliseconds(ExpiresIn); }
        }

        [JsonProperty("bearer")]
        public string Bearer { get; set; }


        public static async Task<Token> GetToken(Uri authenticationUrl, Dictionary<string, string> authenticationCredentials)
        {
            HttpClient client = new HttpClient();

            FormUrlEncodedContent content = new FormUrlEncodedContent(authenticationCredentials);


            //var request = new HttpMessageRequest(authenticationUrl);
            //request.Content = yourContent;
            //var response = await client.SendAsync(request,
            //HttpCompletionOption.ResponseHeadersRead);

            HttpResponseMessage response = await client.PostAsync(authenticationUrl.ToString(), content).ConfigureAwait(false);
           // HttpResponseMessage response = await client.PostAsync(authenticationUrl.ToString(), content);

            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                string message = String.Format("POST failed. Received HTTP {0}", response.StatusCode);
                throw new ApplicationException(message);
            }

            string responseString = await response.Content.ReadAsStringAsync();

            Token token = JsonConvert.DeserializeObject<Token>(responseString);

            return token;
        }

        public async Task<string> PostTest(object sampleData, string uri)
        {
            var request = (HttpWebRequest)WebRequest.Create(new Uri(uri));
            request.ContentType = "application/json";
            request.Method = "POST";
            request.Timeout = 4000; //ms
            var itemToSend = JsonConvert.SerializeObject(sampleData);
            using (var streamWriter = new StreamWriter(await request.GetRequestStreamAsync()))
            {
                streamWriter.Write(itemToSend);
                streamWriter.Flush();
                streamWriter.Dispose();
            }

            // Send the request to the server and wait for the response:  
            using (var response = await request.GetResponseAsync())
            {
                // Get a stream representation of the HTTP web response:  
                using (var stream = response.GetResponseStream())
                {
                    var reader = new StreamReader(stream);
                    var message = reader.ReadToEnd();
                    return message;
                }
            }
        }


    }
}
