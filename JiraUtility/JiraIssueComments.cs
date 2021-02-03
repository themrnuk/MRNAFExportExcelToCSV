using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV.JiraUtility
{
    public class JiraIssueComments
    {
        public string IssueID { get; set; }
        public string IssueKey { get; set; }
        public string CommentID { get; set; }
        public string Author { get; set; }
        public string Body { get; set; }
        public DateTime Created { get; set; }
        public DateTime Updated { get; set; }
    }
}
