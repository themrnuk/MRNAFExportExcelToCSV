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
using Newtonsoft.Json.Linq;
using System.Security;
using Microsoft.SharePoint.Client;
using Microsoft.WindowsAzure.Storage;
using System.Text;
using MRNAFExportExcelToCSV.JiraUtility;
using System.Linq;
using System.Collections.Generic;

namespace MRNAFExportExcelToCSV
{
    public static class JiraAPIData
    {
        [FunctionName("JiraAPIData")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log, ExecutionContext context)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            //string apiName = req.Query["ApiName"];
            //string deltaColumn = req.Query["DeltaColumn"];
            //string modifiedDate = req.Query["ModifiedDate"];

            //string requestBody = await new StreamReader(req.Body).ReadToEndAsync();

            //MSPApi data = JsonConvert.DeserializeObject<MSPApi>(requestBody);
            //apiName = apiName ?? data?.ApiName;
            //deltaColumn = deltaColumn ?? data?.DeltaColumn;
            //modifiedDate = modifiedDate ?? data?.ModifiedDate;

            //if (string.IsNullOrEmpty(modifiedDate))
            //{
            //    return new OkObjectResult("This HTTP triggered function executed successfully. Pass a ModifiedDate in the query string or in the request body for a personalized response.");
            //}

            using (var handler = new HttpClientHandler())
            {
                //Invoke REST API 
                using (var client = new HttpClient(handler))
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    var byteArray = Encoding.ASCII.GetBytes($"{JiraAPICredentials.UserName}:{JiraAPICredentials.AccessToken}");
                    client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                    string url = string.Empty;
                    //======================================Getting Jira Projects============================================
                    string projectJson = "[]";
                    string jiraProjectFileName = string.Format("{0}.json", "JIRA_Projects");
                    url = $"{JiraAPICredentials.APIBaseUrl}/project";
                    HttpResponseMessage projectresponse = await client.GetAsync(url).ConfigureAwait(false);
                    projectresponse.EnsureSuccessStatusCode();
                    string projectJsonData = await projectresponse.Content.ReadAsStringAsync();
                    var projectsettings = new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        MissingMemberHandling = MissingMemberHandling.Ignore
                    };
                    List<JiraProjectResponse> jiraAPIProjectResponse = JsonConvert.DeserializeObject<List<JiraProjectResponse>>(projectJsonData, projectsettings);

                    List<JiraProjects> allprojects = (from JiraProjectResponse project in jiraAPIProjectResponse
                                                 select new JiraProjects
                                                 {
                                                     ProjectID = project.id,
                                                     ProjectName = project.name,
                                                     ProjectKey = project.key,
                                                     IsPrivate = project.isPrivate
                                                 }).ToList();
                    projectJson = JsonConvert.SerializeObject(allprojects);
                    await SaveJSONFileAsync(context, jiraProjectFileName, projectJson);
                    //========================================================================================================


                    //=======================================Getting Jira Issues Pagging Info=========================================
                    url = $"{JiraAPICredentials.APIBaseUrl}/search?fields=total";
                    HttpResponseMessage paggingresponse = await client.GetAsync(url).ConfigureAwait(false);
                    paggingresponse.EnsureSuccessStatusCode();
                    string paggingJsonData = await paggingresponse.Content.ReadAsStringAsync();
                    var paggingsettings = new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        MissingMemberHandling = MissingMemberHandling.Ignore
                    };
                    IssuePaggingDetail jiraAPIPaggingResponse = JsonConvert.DeserializeObject<IssuePaggingDetail>(paggingJsonData, paggingsettings);

                    List<int> pages = new List<int>();
                    if (jiraAPIPaggingResponse != null && jiraAPIPaggingResponse.total >= 100)
                    {
                        int lastPage = jiraAPIPaggingResponse.total / 100;
                        for (int i = 0; i <= lastPage; i++)
                        {
                            pages.Add(i * 100);
                        }
                    }
                    //==================================================================================================================


                    List<JiraIssues> allissues = new List<JiraIssues>();
                    List<JiraSprints> allsprints = new List<JiraSprints>();
                    List<JiraIssueLabels> allissuelabels = new List<JiraIssueLabels>();
                    List<IssueComponents> allissuecomponents = new List<IssueComponents>();
                    List<JiraIssueComments> allissuecomments = new List<JiraIssueComments>();
                    List<JiraIssueWorkLogs> allissueworklogs = new List<JiraIssueWorkLogs>();
                    List<JiraIssueHistories> allissuehistories = new List<JiraIssueHistories>();

                    string issuesJson = "[]";
                    string sprintJson = "[]";
                    string labelJson = "[]";
                    string componentJson = "[]";
                    string commmentJson = "[]";
                    string worklogJson = "[]";
                    string historyJson = "[]";
                    foreach (int page in pages)
                    {
                        url = $"{JiraAPICredentials.APIBaseUrl}/search?fields=comment,worklog,summary,description,customfield_10020,customfield_10024,assignee,creator,reporter,priority,project,timetracking,labels,components,created,updated,status&expand=changelog&maxResults=100&startAt=" + page;
                        HttpResponseMessage response = await client.GetAsync(url).ConfigureAwait(false);
                        response.EnsureSuccessStatusCode();
                        string jsonData = await response.Content.ReadAsStringAsync();
                        var settings = new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore,
                            MissingMemberHandling = MissingMemberHandling.Ignore
                        };
                        JiraAPIResponse jiraAPIResponse = JsonConvert.DeserializeObject<JiraAPIResponse>(jsonData, settings);

                        if (jiraAPIResponse != null)
                        {
                            //===============================Getting Issues===============================================================
                            List<JiraIssues> issues = (from Issue issue in jiraAPIResponse.issues
                                                       select new JiraIssues
                                                       {
                                                           IssueID = issue.id,
                                                           IssueKey = issue.key,
                                                           Assignee = issue.fields.assignee != null ? issue.fields.assignee.displayName : "",
                                                           Created = issue.fields.created,
                                                           Creator = issue.fields.creator != null ? issue.fields.creator.displayName : "",
                                                           Description = issue.fields.description,
                                                           Priority = issue.fields.priority != null ? issue.fields.priority.name : "",
                                                           ProjectID = issue.fields.project != null ? issue.fields.project.id : "",
                                                           Project = issue.fields.project != null ? issue.fields.project.name : "",
                                                           RemainingEstimate = issue.fields.timetracking != null ? issue.fields.timetracking.remainingEstimate : "",
                                                           RemainingEstimateSeconds = issue.fields.timetracking != null ? issue.fields.timetracking.remainingEstimateSeconds : 0,
                                                           TimeSpent = issue.fields.timetracking != null ? issue.fields.timetracking.timeSpent : "",
                                                           TimeSpentSeconds = issue.fields.timetracking != null ? issue.fields.timetracking.timeSpentSeconds : 0,
                                                           Reporter = issue.fields.reporter != null ? issue.fields.reporter.displayName : "",
                                                           Status = issue.fields.status != null ? issue.fields.status.name : "",
                                                           StoryPoints = issue.fields.customfield_10024,
                                                           Summary = issue.fields.summary,
                                                           Updated = issue.fields.updated,
                                                       }).ToList();
                            allissues.AddRange(issues);
                            issuesJson = JsonConvert.SerializeObject(allissues);
                            //==============================================================================================================

                            //=======================================Getting Sprints========================================================
                            var issuesprints = (from Issue issue in jiraAPIResponse.issues
                                                select new
                                                {
                                                    IssueID = issue.id,
                                                    IssueKey = issue.key,
                                                    customfield10020 = issue.fields.customfield_10020
                                                }).ToList();

                            foreach (var issue in issuesprints)
                            {
                                if (issue.customfield10020 != null)
                                {
                                    foreach (var sprint in issue.customfield10020)
                                    {
                                        JiraSprints jiraSprint = new JiraSprints();
                                        jiraSprint.IssueID = Convert.ToInt32(issue.IssueID);
                                        jiraSprint.IssueKey = issue.IssueKey;
                                        if (sprint != null)
                                        {
                                            jiraSprint.EndDate = sprint.endDate;
                                            jiraSprint.Name = sprint.name;
                                            jiraSprint.SprintID = sprint.id;
                                            jiraSprint.StartDate = sprint.startDate;
                                            jiraSprint.State = sprint.state;
                                        }

                                        allsprints.Add(jiraSprint);
                                    }
                                }

                            }
                            sprintJson = JsonConvert.SerializeObject(allsprints);
                            //================================================================================================================

                            //=========================================Getting Issue Labels===================================================
                            var issueLabels = (from Issue issue in jiraAPIResponse.issues
                                               select new
                                               {
                                                   IssueID = issue.id,
                                                   IssueKey = issue.key,
                                                   labels = issue.fields.labels
                                               }).ToList();

                            foreach (var issue in issueLabels)
                            {
                                if (issue.labels != null)
                                {
                                    foreach (var label in issue.labels)
                                    {
                                        JiraIssueLabels jiraIssueLabel = new JiraIssueLabels();
                                        jiraIssueLabel.IssueID = Convert.ToInt32(issue.IssueID);
                                        jiraIssueLabel.IssueKey = issue.IssueKey;
                                        jiraIssueLabel.Value = Convert.ToString(label);
                                        allissuelabels.Add(jiraIssueLabel);
                                    }
                                }
                            }
                            labelJson = JsonConvert.SerializeObject(allissuelabels);
                            //================================================================================================================

                            //==========================================Getting Issue Components==============================================
                            var issuecomonents = (from Issue issue in jiraAPIResponse.issues
                                                  select new
                                                  {
                                                      IssueID = issue.id,
                                                      IssueKey = issue.key,
                                                      components = issue.fields.components
                                                  }).ToList();

                            foreach (var issue in issuecomonents)
                            {
                                if (issue.components != null)
                                {
                                    foreach (var component in issue.components)
                                    {
                                        IssueComponents jiraIssueComponent = new IssueComponents();
                                        jiraIssueComponent.IssueID = issue.IssueID;
                                        jiraIssueComponent.IssueKey = issue.IssueKey;
                                        if (component != null)
                                        {
                                            jiraIssueComponent.ComponentID = component.id;
                                            jiraIssueComponent.Name = component.name;
                                        }
                                        allissuecomponents.Add(jiraIssueComponent);
                                    }
                                }
                            }
                            componentJson = JsonConvert.SerializeObject(allissuecomponents);
                            //================================================================================================================

                            //=====================================Getting Issue Comments==============================================================
                            var issuecomments = (from Issue issue in jiraAPIResponse.issues
                                                 select new
                                                 {
                                                     IssueID = issue.id,
                                                     IssueKey = issue.key,
                                                     comments = issue.fields.comment != null ? issue.fields.comment.comments : new List<Comment2>()
                                                 }).ToList();
                            foreach (var issue in issuecomments)
                            {
                                if (issue.comments != null)
                                {
                                    foreach (var comment in issue.comments)
                                    {
                                        JiraIssueComments jiraIssueComment = new JiraIssueComments();
                                        jiraIssueComment.IssueID = issue.IssueID;
                                        jiraIssueComment.IssueKey = issue.IssueKey;
                                        if (comment != null)
                                        {
                                            jiraIssueComment.CommentID = comment.id;
                                            jiraIssueComment.Author = comment.author.displayName;
                                            jiraIssueComment.Body = comment.body;
                                            jiraIssueComment.Created = comment.created;
                                            jiraIssueComment.Updated = comment.updated;
                                        }
                                        allissuecomments.Add(jiraIssueComment);
                                    }
                                }
                            }
                            commmentJson = JsonConvert.SerializeObject(allissuecomments);


                            //=================================================================================================================

                            //=====================================Grtting Issue Histories=============================================================
                            var issuehistories = (from Issue issue in jiraAPIResponse.issues
                                                  select new
                                                  {
                                                      IssueID = issue.id,
                                                      IssueKey = issue.key,
                                                      histories = issue.changelog != null ? issue.changelog.histories : new List<History>()
                                                  }).ToList();
                            foreach (var issue in issuehistories)
                            {
                                if (issue != null && issue.histories != null)
                                {
                                    foreach (var history in issue.histories)
                                    {
                                        if (history != null)
                                        {
                                            foreach (var item in history.items)
                                            {
                                                if (item != null)
                                                {
                                                    JiraIssueHistories jiraIssuehistory = new JiraIssueHistories();
                                                    jiraIssuehistory.HistoryID = history.id;
                                                    jiraIssuehistory.Auther = history.author.displayName;
                                                    jiraIssuehistory.Created = history.created;
                                                    jiraIssuehistory.IssueID = issue.IssueID;
                                                    jiraIssuehistory.IssueKey = issue.IssueKey;
                                                    jiraIssuehistory.Field = item.field;
                                                    jiraIssuehistory.Fieldtype = item.fieldtype;
                                                    jiraIssuehistory.FieldId = item.fieldId;
                                                    jiraIssuehistory.From = item.from;
                                                    jiraIssuehistory.FromString = item.fromString;
                                                    jiraIssuehistory.To = item.to;
                                                    jiraIssuehistory.Tostring = item.toString;
                                                    allissuehistories.Add(jiraIssuehistory);
                                                }
                                            }
                                        }

                                    }
                                }
                            }
                            historyJson = JsonConvert.SerializeObject(allissuehistories);
                            //=================================================================================================================

                            //=====================================Getting Issue Worklogs==============================================================

                            var issueworklogs = (from Issue issue in jiraAPIResponse.issues
                                                 select new
                                                 {
                                                     IssueID = issue.id,
                                                     IssueKey = issue.key,
                                                     worklogs = issue.fields.worklog != null ? issue.fields.worklog.worklogs : new List<Worklog2>()
                                                 }).ToList();
                            foreach (var issue in issueworklogs)
                            {
                                if (issue.worklogs != null)
                                {
                                    foreach (var worklog in issue.worklogs)
                                    {
                                        JiraIssueWorkLogs jiraIssueworklog = new JiraIssueWorkLogs();
                                        jiraIssueworklog.IssueID = issue.IssueID;
                                        jiraIssueworklog.IssueKey = issue.IssueKey;
                                        if (worklog != null)
                                        {
                                            jiraIssueworklog.WorklogID = worklog.id;
                                            jiraIssueworklog.Author = worklog.author.displayName;
                                            jiraIssueworklog.Comment = worklog.author.displayName; ;
                                            jiraIssueworklog.Created = worklog.created;
                                            jiraIssueworklog.Updated = worklog.updated;
                                            jiraIssueworklog.Started = worklog.started;
                                            jiraIssueworklog.TimeSpent = worklog.timeSpent; ;
                                            jiraIssueworklog.TimeSpentSeconds = worklog.timeSpentSeconds;
                                        }
                                        allissueworklogs.Add(jiraIssueworklog);
                                    }
                                }
                            }
                            worklogJson = JsonConvert.SerializeObject(allissueworklogs);
                            //=================================================================================================================
                        }
                    }

                    string jiraIssueFileName = string.Format("{0}.json", "JIRA_Issues");
                    string jiraSprintFileName = string.Format("{0}.json", "JIRA_Sprints");
                    string jiraIssueLabelsFileName = string.Format("{0}.json", "JIRA_Issue_Labels");
                    string jiraIssueComponentFileName = string.Format("{0}.json", "JIRA_Issue_Components");
                    string jiraIssueCommentFileName = string.Format("{0}.json", "JIRA_Issue_Comments");
                    string jiraIssueHistoryFileName = string.Format("{0}.json", "JIRA_Issue_History");
                    string jiraIssueWorklogFileName = string.Format("{0}.json", "JIRA_Issue_Worklogs");

                    await SaveJSONFileAsync(context, jiraIssueFileName, issuesJson);
                    await SaveJSONFileAsync(context, jiraSprintFileName, sprintJson);
                    await SaveJSONFileAsync(context, jiraIssueLabelsFileName, labelJson);
                    await SaveJSONFileAsync(context, jiraIssueComponentFileName, componentJson);
                    await SaveJSONFileAsync(context, jiraIssueCommentFileName, commmentJson);
                    await SaveJSONFileAsync(context, jiraIssueHistoryFileName, historyJson);
                    await SaveJSONFileAsync(context, jiraIssueWorklogFileName, worklogJson);

                    JiraFunctionResponse jiraFunctionResponse = new JiraFunctionResponse();
                    jiraFunctionResponse.IsSuccess = true;
                    jiraFunctionResponse.Message = "Jira data imported successfully.";
                    return new OkObjectResult(jiraFunctionResponse);
                }
            }
        }

        public static async Task SaveJSONFileAsync(ExecutionContext context, string filename, string issuesJson)
        {
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
            var destBlob = container.GetBlockBlobReference($"Jira/{filename}");
            await destBlob.UploadTextAsync(issuesJson);
        }

    }
}
