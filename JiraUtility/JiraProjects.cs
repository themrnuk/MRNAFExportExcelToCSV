using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV.JiraUtility
{
    public class JiraProjects
    {
        public string ProjectID { get; set; }
        public string ProjectKey { get; set; }
        public string ProjectName { get; set; }
        public bool IsPrivate { get; set; }
    }
}
