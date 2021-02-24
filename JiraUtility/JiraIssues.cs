using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV.JiraUtility
{
    public class JiraIssues
    {
        public string IssueID { get; set; }
        public string IssueKey { get; set; }
        public string Summary { get; set; }
        public string Creator { get; set; }
        public DateTime Created { get; set; }
        public string Description { get; set; }

        public string ProjectID { get; set; }
        public string Project { get; set; }
        public string Reporter { get; set; }
        public string Priority { get; set; }
        public double StoryPointEstimate { get; set; }
        public double StoryPoints { get; set; }
        public string RemainingEstimate { get; set; }
        public string TimeSpent { get; set; }
        public int RemainingEstimateSeconds { get; set; }
        public int TimeSpentSeconds { get; set; }
        public string Assignee { get; set; }
        public DateTime Updated { get; set; }
        public string Status { get; set; }
    }
}
