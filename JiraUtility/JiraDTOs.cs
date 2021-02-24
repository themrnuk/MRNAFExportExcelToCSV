using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV.JiraUtility
{
    public class JiraTestCase
    {
        public int TestCaseID { get; set; }
        public string TestCaseKey { get; set; }
        public string Name { get; set; }
        public int ProjectID { get; set; }
        public DateTime CreatedOn { get; set; }
        public string Objective { get; set; }
        public string Precondition { get; set; }
        public int EstimatedTime { get; set; }
        public int ComponentID { get; set; }
        public int PriorityID { get; set; }
        public int StatusID { get; set; }
        public int FolderID { get; set; }
        public string OwnerAccountID { get; set; }
    }
    public class JiraTestCaseStatus
    {
        public int StatusID { get; set; }
        public int ProjectID { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public int Index { get; set; }
        public string Color { get; set; }
        public bool Archived { get; set; }
        public bool Default { get; set; }
    }

    public class JiraTestCaseProject
    {
        public int ProjectID { get; set; }
        public int JiraProjectID { get; set; }
        public string Key { get; set; }
        public bool Enabled { get; set; }
    }

    public class JiraTestCasePriority
    {
        public int PriorityID { get; set; }
        public int ProjectID { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public int Index { get; set; }
        public bool Default { get; set; }
    }

    public class JiraTestCaseExecution
    {
        public int ExecutionID { get; set; }
        public string Key { get; set; }
        public int ProjectID { get; set; }
        public int TestCaseID { get; set; }
        public int EnvironmentID { get; set; }
        public int StatusID { get; set; }
        public DateTime ActualEndDate { get; set; }
        public int EstimatedTime { get; set; }
        public int ExecutionTime { get; set; }
        public string ExecutedByID { get; set; }
        public string AssignedToID { get; set; }
        public string Comment { get; set; }
        public bool Automated { get; set; }
        public int TestCycleID { get; set; }
    }

    public class JiraTestCaseCycle
    {
        public int CycleID { get; set; }
        public string Key { get; set; }
        public string Name { get; set; }
        public int ProjectID { get; set; }
        public int StatusID { get; set; }
        public string Description { get; set; }
        public DateTime PlannedStartDate { get; set; }
        public DateTime PlannedEndDate { get; set; }
        public string OwnerAccountID { get; set; }
    }

    public class JiraTestCaseEnvironment
    {
        public int EnvironmentID { get; set; }
        public int ProjectID { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public int Index { get; set; }
        public bool Archived { get; set; }
    }


}
