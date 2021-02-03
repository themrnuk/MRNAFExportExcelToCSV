using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV.JiraUtility
{
  public  class JiraIssueWorkLogs
    {
        public string IssueID { get; set; }
        public string IssueKey { get; set; }
        public string Author { get; set; }
        public string Comment { get; set; }
        public DateTime Created { get; set; }
        public DateTime Updated { get; set; }
        public DateTime Started { get; set; }
        public string TimeSpent { get; set; }
        public int TimeSpentSeconds { get; set; }
        public string WorklogID { get; set; }
       
    }
}
