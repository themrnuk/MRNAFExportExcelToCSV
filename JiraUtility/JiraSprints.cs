using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV.JiraUtility
{
   public class JiraSprints
    {
        public int IssueID { get; set; }
        public string IssueKey { get; set; }
        public int SprintID { get; set; }
        public string Name { get; set; }
        public string State { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
    }
}
