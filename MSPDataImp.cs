using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.Azure.Storage;
using Microsoft.Azure.Storage.Blob;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using AngleSharp.Dom.Css;
using System.Collections.Generic;

namespace MRNAFExportExcelToCSV
{
   
    //public class ProjectData
    //{
    //    public string name { get; set; }
    //    public string url { get; set; }
    //}
    public static class MSPDataImp
    {
        //private static string baseUrl="https://themrn.sharepoint.com";
        //private static string userName = "mahavir.rawat@themrn.co.uk";
        //private static string password = "HeavenlyHappy270420£";
         
        public static async Task<string> GetProjectData()
        {                       
            var credential = MSPCredential.Credentials ;
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

        public static async Task SaveCsv(string projectData, ExecutionContext context)
        {
            var securePassword = new SecureString();
            string url = $"https://themrn.sharepoint.com/sites/pwa/_api/ProjectData/[en-us]/{projectData}?format=json";
            var credential = MSPCredential.Credentials;
            using (var handler = new HttpClientHandler() { Credentials = credential })
            {
                Uri uri = new Uri(MSPCredential.BaseUrl);
                handler.CookieContainer.SetCookies(uri, credential.GetAuthenticationCookie(uri));

                //Invoke REST API 
                using (var client = new HttpClient(handler))
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    HttpResponseMessage response = await client.GetAsync(url).ConfigureAwait(false);
                    response.EnsureSuccessStatusCode();

                    string jsonData = await response.Content.ReadAsStringAsync();                                    
                    string jsonFile = string.Format("{0}.json",  projectData);
                    var config = new ConfigurationBuilder()
                  .SetBasePath(context.FunctionAppDirectory)
                  .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
                  .AddEnvironmentVariables()
                  .Build();
                    string azurestorageconnectionString = config.GetConnectionStringOrSetting("AzureWebJobsStorage");
                    CloudStorageAccount storageAccount = CloudStorageAccount.Parse(azurestorageconnectionString);

                    var blobClient = storageAccount.CreateCloudBlobClient();
                    var container = blobClient.GetContainerReference("importfiles");
                    var destBlob = container.GetBlockBlobReference($"MSP/{jsonFile}");
                    destBlob.UploadText(JObject.Parse(jsonData)["value"].ToString());                    
                }
            }
        }
        [FunctionName("MSPDataImp")]
        public static void Run([TimerTrigger("0 10 12 * * *")] TimerInfo myTimer, ILogger log, ExecutionContext context)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
            
            var data = GetProjectData();            
            Task.WaitAll(data);
             JObject rss = JObject.Parse(data.Result);
            JArray mspArray = (JArray)rss["value"];
            var projectDatas = mspArray.ToObject<List<ProjectData>>();
            List<Task> tasksToWait = new List<Task>();   

            List<Task> taskArray = new List<Task>();
            foreach (var projectData in projectDatas)
            {
                taskArray.Add(SaveCsv(projectData.url, context));
            }
            Task.WaitAll(taskArray.ToArray());           
        }
    }
}
