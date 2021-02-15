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
//using Microsoft.Azure.Storage;
using Newtonsoft.Json.Linq;
//using Microsoft.Azure.Storage.Blob;
using System.Security;
using Microsoft.SharePoint.Client;
using Microsoft.WindowsAzure.Storage;
using RestSharp;
//using Microsoft.Graph;

namespace MRNAFExportExcelToCSV
{

    public static class MSPApiData
    {
        
        #region MSPApiData        

        [FunctionName("MSPApiData")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log, ExecutionContext context)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string apiName = req.Query["ApiName"];
            string deltaColumn = req.Query["DeltaColumn"];
            string modifiedDate = req.Query["ModifiedDate"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();

            MSPApi data = JsonConvert.DeserializeObject<MSPApi>(requestBody);
            apiName = apiName ?? data?.ApiName;
            deltaColumn = deltaColumn ?? data?.DeltaColumn;
            modifiedDate = modifiedDate ?? data?.ModifiedDate;

            if (string.IsNullOrEmpty(apiName))
            {
                return new OkObjectResult("This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.");
            }

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
                    string url = string.Empty;
                    if (string.IsNullOrEmpty(deltaColumn))
                    {
                        url = $"{MSPCredential.BaseUrl}/sites/pwa/_api/ProjectData/[en-us]/{apiName}?format=json";
                    }
                    else
                    {
                        url = $"{MSPCredential.BaseUrl}/sites/pwa/_api/ProjectData/[en-us]/{apiName}?format=json&$Filter={deltaColumn} eq null or {deltaColumn} ge datetime'{modifiedDate}'";
                    }

                    log.LogInformation($"API Url is - {url}");

                    HttpResponseMessage response = await client.GetAsync(url).ConfigureAwait(false);
                    response.EnsureSuccessStatusCode();

                    string jsonData = await response.Content.ReadAsStringAsync();
                    string jsonFile = string.Format("{0}.json", apiName);
                    var config = new ConfigurationBuilder()
                  .SetBasePath(context.FunctionAppDirectory)
                  .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
                  .AddEnvironmentVariables()
                  .Build();
                    string defaultContainerName = config.GetConnectionStringOrSetting("ContainerName");
                    string azurestorageconnectionString = config.GetConnectionStringOrSetting("AzureWebJobsStorage");
                    CloudStorageAccount storageAccount = CloudStorageAccount.Parse(azurestorageconnectionString);

                    var blobClient = storageAccount.CreateCloudBlobClient();
                    var container = blobClient.GetContainerReference(defaultContainerName);
                    var destBlob = container.GetBlockBlobReference($"MSP/{jsonFile}");
                    await destBlob.UploadTextAsync(JObject.Parse(jsonData)["value"].ToString());
                    MSPResponseDataApi mspResponseDataApi = new MSPResponseDataApi();
                    mspResponseDataApi.JsonFileName = jsonFile;
                    return new OkObjectResult(mspResponseDataApi);
                }
            }
        }

        #endregion

        //[FunctionName("MSPApiData")]
        //public static async Task<IActionResult> Run(
        //    [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
        //    ILogger log, ExecutionContext context)
        //{
        //    log.LogInformation("C# HTTP trigger function processed a request.");

        //    string apiName = req.Query["ApiName"];
        //    string deltaColumn = req.Query["DeltaColumn"];
        //    string modifiedDate = req.Query["ModifiedDate"];

        //    string requestBody = await new StreamReader(req.Body).ReadToEndAsync();

        //    MSPApi data = JsonConvert.DeserializeObject<MSPApi>(requestBody);
        //    apiName = apiName ?? data?.ApiName;
        //    deltaColumn = deltaColumn ?? data?.DeltaColumn;
        //    modifiedDate = modifiedDate ?? data?.ModifiedDate;

        //    if (string.IsNullOrEmpty(apiName))
        //    {
        //        return new OkObjectResult("This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.");
        //    }

        //    //var credential = MSPCredential.Credentials;
        //    //using (var handler = new HttpClientHandler() { Credentials = credential })
        //    //{

        //    //    //Get authentication cookie
        //    //    Uri uri = new Uri(MSPCredential.BaseUrl);
        //    //    handler.CookieContainer.SetCookies(uri, credential.GetAuthenticationCookie(uri));

        //    //    //Invoke REST API 
        //    //    using (var client = new HttpClient(handler))
        //    //    {
        //    //        client.DefaultRequestHeaders.Accept.Clear();
        //    //        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        //    //        string url = string.Empty;
        //    //        if (string.IsNullOrEmpty(deltaColumn))
        //    //        {
        //    //            url = $"{MSPCredential.BaseUrl}/sites/pwa/_api/ProjectData/[en-us]/{apiName}?format=json";
        //    //        }
        //    //        else
        //    //        {
        //    //            url = $"{MSPCredential.BaseUrl}/sites/pwa/_api/ProjectData/[en-us]/{apiName}?format=json&$Filter={deltaColumn} eq null or {deltaColumn} ge datetime'{modifiedDate}'";
        //    //        }                    

        //    //        HttpResponseMessage response = await client.GetAsync(url).ConfigureAwait(false);
        //    //        response.EnsureSuccessStatusCode();

        //    //        string jsonData = await response.Content.ReadAsStringAsync();
        //    //        string jsonFile = string.Format("{0}.json", apiName);
        //    //        var config = new ConfigurationBuilder()
        //    //      .SetBasePath(context.FunctionAppDirectory)
        //    //      .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
        //    //      .AddEnvironmentVariables()
        //    //      .Build();
        //    //        string defaultContainerName = config.GetConnectionStringOrSetting("ContainerName");
        //    //        string azurestorageconnectionString = config.GetConnectionStringOrSetting("AzureWebJobsStorage");
        //    //        CloudStorageAccount storageAccount = CloudStorageAccount.Parse(azurestorageconnectionString);

        //    //        var blobClient = storageAccount.CreateCloudBlobClient();
        //    //        var container = blobClient.GetContainerReference(defaultContainerName);
        //    //        var destBlob = container.GetBlockBlobReference($"MSP/{jsonFile}");
        //    //        destBlob.UploadText(JObject.Parse(jsonData)["value"].ToString());                   
        //    //        MSPResponseDataApi mspResponseDataApi = new MSPResponseDataApi();
        //    //        mspResponseDataApi.JsonFileName = jsonFile;                    
        //    //        return new OkObjectResult(mspResponseDataApi);
        //    //    }
        //    //}

        //    #region URL

        //    string stGetAccessTokenUrl = "https://accounts.accesscontrol.windows.net/{0}/tokens/OAuth/2";
        //    string tenantID = "0bb4595f-62c0-4464-82ce-71670c150b3b";
        //    string resourceID = "00000003-0000-0ff1-ce00-000000000000";
        //    string appClientID = "b7a482a3-8437-46d8-8bd9-37bf5b647a6b";
        //    string appClientSecret = "W7~hmY2-UU8tBwzQAEf2O-tTB-Wi~Y0ToR";
        //    string appSiteDomain = "themrn.sharepoint.com";


        //    string appSiteDetailsUrl = string.Empty;
        //    if (string.IsNullOrEmpty(deltaColumn))
        //    {
        //        appSiteDetailsUrl = $"{MSPCredential.BaseUrl}/sites/pwa/_api/ProjectData/[en-us]/{apiName}?format=json";
        //    }
        //    else
        //    {
        //        appSiteDetailsUrl = $"{MSPCredential.BaseUrl}/sites/pwa/_api/ProjectData/[en-us]/{apiName}?format=json&$Filter={deltaColumn} eq null or {deltaColumn} ge datetime'{modifiedDate}'";
        //    }


        //    #endregion

        //    string accessToken = string.Empty;
        //    var apiClientToGetToken = new RestClient(string.Format(stGetAccessTokenUrl, tenantID));
        //    apiClientToGetToken.Timeout = -1;
        //    var requestToGetToken = new RestRequest(Method.POST);
        //    requestToGetToken.AddParameter("grant_type", "client_credentials",RestSharp.ParameterType.GetOrPost);
        //    requestToGetToken.AddParameter("client_id", appClientID + "@" + tenantID, RestSharp.ParameterType.GetOrPost);
        //    requestToGetToken.AddParameter("client_secret", appClientSecret, RestSharp.ParameterType.GetOrPost);
        //    requestToGetToken.AddParameter("resource", resourceID + "/" + appSiteDomain + "@" + tenantID, RestSharp.ParameterType.GetOrPost);
        //    IRestResponse tokenResponse = apiClientToGetToken.Execute(requestToGetToken);

        //    string[] tokenResponseKeys = tokenResponse.Content.Split(',');

        //    foreach (var key in tokenResponseKeys)
        //    {
        //        if (key.StartsWith("\"access_token\":"))
        //        {
        //            accessToken = key.Substring(16);
        //            accessToken = accessToken.Substring(0, accessToken.Length - 2);
        //        }
        //    }

        //    var clientToGetData = new RestClient(appSiteDetailsUrl);
        //    clientToGetData.Timeout = -1;
        //    var requestToGetData = new RestRequest(Method.GET);
        //    requestToGetData.AddHeader("Authorization", "Bearer " + accessToken);
        //    requestToGetData.AddHeader("Accept", "application/json; odata=verbose");
        //    IRestResponse dataResponse = clientToGetData.Execute(requestToGetData);


        //    string jsonData = dataResponse.Content;
        //    string jsonFile = string.Format("{0}.json", apiName);
        //    var config = new ConfigurationBuilder()
        //  .SetBasePath(context.FunctionAppDirectory)
        //  .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
        //  .AddEnvironmentVariables()
        //  .Build();
        //    string defaultContainerName = config.GetConnectionStringOrSetting("ContainerName");
        //    string azurestorageconnectionString = config.GetConnectionStringOrSetting("AzureWebJobsStorage");
        //    CloudStorageAccount storageAccount = CloudStorageAccount.Parse(azurestorageconnectionString);

        //    var blobClient = storageAccount.CreateCloudBlobClient();
        //    var container = blobClient.GetContainerReference(defaultContainerName);
        //    var destBlob = container.GetBlockBlobReference($"MSP/{jsonFile}");
        //    await destBlob.UploadTextAsync(jsonData);
        //    //destBlob.UploadText(JObject.Parse(jsonData)["value"].ToString()); MSPResponseDataApi mspResponseDataApi = new MSPResponseDataApi();
        //    MSPResponseDataApi mspResponseDataApi = new MSPResponseDataApi();
        //    mspResponseDataApi.JsonFileName = jsonFile;
        //    return new OkObjectResult(mspResponseDataApi);
        //}
    }
}
