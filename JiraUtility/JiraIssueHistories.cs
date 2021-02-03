using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV.JiraUtility
{
    public class JiraIssueHistories
    {
        public string HistoryID { get; set; }
        public string Auther { get; set; }
        public DateTime Created { get; set; }
        public string IssueID { get; set; }
        public string IssueKey { get; set; }
        public string Field { get; set; }
        public string Fieldtype { get; set; }
        public string FieldId { get; set; }
        public string From { get; set; }
        public string FromString { get; set; }
        public string To { get; set; }
        public string Tostring { get; set; }
    }
}
