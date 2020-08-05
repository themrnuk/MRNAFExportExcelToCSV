using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Net.Http;
using System.Net.Http.Headers;
using Microsoft.Extensions.Configuration;
using Microsoft.Azure.Storage;
using Newtonsoft.Json.Linq;
using Microsoft.Azure.Storage.Blob;
using System.Security;
using Microsoft.SharePoint.Client;
//using Microsoft.Graph;

namespace MRNAFExportExcelToCSV
{
    public static class MSPApiList
    {
        #region MSPApiList
        public static async Task<string> GetProjectData()
        {
            var credential = MSPCredential.Credentials;
            using (var handler = new HttpClientHandler() { Credentials = credential })
            {
                //Get authentication cookie
                Uri uri = new Uri(MSPCredential.BaseUrl);
                handler.CookieContainer.SetCookies(uri, credential.GetAuthenticationCookie(uri));

                //Invoke REST API 
                using (var client = new HttpClient(handler))
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    string dataUrl = string.Format("{0}/sites/pwa/_api/ProjectData/[en-us]", MSPCredential.BaseUrl);
                    HttpResponseMessage response = await client.GetAsync(dataUrl).ConfigureAwait(false);
                    response.EnsureSuccessStatusCode();

                    string jsonData = await response.Content.ReadAsStringAsync();
                    return jsonData;

                }
            }

        }

        [FunctionName("MSPApiList")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            var data = GetProjectData();
            Task.WaitAll(data);
            JObject rss = JObject.Parse(data.Result);
            JArray mspArray = (JArray)rss["value"];
            
            var testData=     JsonConvert.SerializeObject(mspArray, Formatting.Indented);
            return new OkObjectResult(testData);
        }

        #endregion
    }
}
