using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV.JiraUtility
{

    public class TestProject
    {
        public int id { get; set; }
        public string self { get; set; }
    }

    public class Component1
    {
        public int id { get; set; }
        public string self { get; set; }
    }

    public class Priority1
    {
        public int id { get; set; }
        public string self { get; set; }
    }

    public class Status1
    {
        public int id { get; set; }
        public string self { get; set; }
    }

    public class Folder
    {
        public int id { get; set; }
        public string self { get; set; }
    }

    public class Owner
    {
        public string self { get; set; }
        public string accountId { get; set; }
    }

    public class TestScript
    {
        public string self { get; set; }
    }

    public class CustomFields1
    {
        public int BuildNumber { get; set; }
        public string ReleaseDate { get; set; }
        public bool Implemented { get; set; }
        public List<string> Category { get; set; }
        public string Tester { get; set; }
    }

    public class Value
    {
        public int id { get; set; }
        public string key { get; set; }
        public string name { get; set; }
        public TestProject project { get; set; }
        public DateTime createdOn { get; set; }
        public string objective { get; set; }
        public string precondition { get; set; }
        public int estimatedTime { get; set; }
        public List<string> labels { get; set; }
        public Component1 component { get; set; }
        public Priority1 priority { get; set; }
        public Status1 status { get; set; }
        public Folder folder { get; set; }
        public Owner owner { get; set; }
        public TestScript testScript { get; set; }
        public CustomFields1 customFields { get; set; }
    }

    public class JiraTestcaseResponse
    {
        public string next { get; set; }
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public int total { get; set; }
        public bool isLast { get; set; }
        public List<Value> values { get; set; }
    }

    public class TestCaseProject
    {
        public int id { get; set; }
        public int jiraProjectId { get; set; }
        public string key { get; set; }
        public bool enabled { get; set; }
    }

    public class TestCaseProjectResponse
    {
        public string next { get; set; }
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public int total { get; set; }
        public bool isLast { get; set; }
        public List<TestCaseProject> values { get; set; }
    }

  


    public class TestCasePriority
    {
        public int id { get; set; }
        public TestProject project { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public int index { get; set; }
        public bool @default { get; set; }
    }

    public class JiraTestCasePriorityResponse
    {
        public string next { get; set; }
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public int total { get; set; }
        public bool isLast { get; set; }
        public List<TestCasePriority> values { get; set; }
    }

   


   
    public class TestCase
    {
        public int id { get; set; }
        public string self { get; set; }
        public int name { get; set; }
    }

    public class TestExecutionStatus
    {
        public int id { get; set; }
        public string self { get; set; }
    }

    public class TestCycle
    {
        public int id { get; set; }
        public string self { get; set; }
    }

    public class TestCaseExecution
    {
        public int id { get; set; }
        public string key { get; set; }
        public TestProject project { get; set; }
        public TestCase testCase { get; set; }
        public TestCaseEnvironment environment { get; set; }
        public TestExecutionStatus testExecutionStatus { get; set; }
        public DateTime actualEndDate { get; set; }
        public int estimatedTime { get; set; }
        public int executionTime { get; set; }
        public string executedById { get; set; }
        public string assignedToId { get; set; }
        public string comment { get; set; }
        public bool automated { get; set; }
        public TestCycle testCycle { get; set; }
    }
    public class JiraTestCaseExecutionResponse
    {
        public string next { get; set; }
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public int total { get; set; }
        public bool isLast { get; set; }
        public List<TestCaseExecution> values { get; set; }
    }



    public class TestCaseCycle
    {
        public int id { get; set; }
        public string key { get; set; }
        public string name { get; set; }
        public TestProject project { get; set; }
        public Status1 status { get; set; }
        public Folder folder { get; set; }
        public string description { get; set; }
        public DateTime plannedStartDate { get; set; }
        public DateTime plannedEndDate { get; set; }
        public Owner owner { get; set; }
    }

    public class JiraTestCaseCycleResponse
    {
        public string next { get; set; }
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public int total { get; set; }
        public bool isLast { get; set; }
        public List<TestCaseCycle> values { get; set; }
    }


    public class TestCaseStatus
    {
        public int id { get; set; }
        public TestProject project { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public int index { get; set; }
        public string color { get; set; }
        public bool archived { get; set; }
        public bool @default { get; set; }
    }

    public class JiraTestCaseStatuesResponse
    {
        public string next { get; set; }
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public int total { get; set; }
        public bool isLast { get; set; }
        public List<TestCaseStatus> values { get; set; }
    }

    public class TestCaseEnvironment
    {
        public int id { get; set; }
        public TestProject project { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public int index { get; set; }
        public bool archived { get; set; }
    }

    public class TestCaseEnvironmentResponse
    {
        public string next { get; set; }
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public int total { get; set; }
        public bool isLast { get; set; }
        public List<TestCaseEnvironment> values { get; set; }
    }



}
