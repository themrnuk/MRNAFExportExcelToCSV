using System;
using System.Collections.Generic;
using System.Text;

namespace MRNAFExportExcelToCSV.JiraUtility
{

    public class JiraAPIResponse
    {
        public string expand { get; set; }
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public int total { get; set; }
        public List<Issue> issues { get; set; }
    }
    public class AvatarUrls
    {
        public string _48x48 { get; set; }
        public string _24x24 { get; set; }
        public string _16x16 { get; set; }
        public string _32x32 { get; set; }
    }

    public class Author
    {
        public string self { get; set; }
        public string accountId { get; set; }
        public AvatarUrls avatarUrls { get; set; }
        public string displayName { get; set; }
        public bool active { get; set; }
        public string timeZone { get; set; }
        public string accountType { get; set; }
    }

    public class Item
    {
        public string field { get; set; }
        public string fieldtype { get; set; }
        public string fieldId { get; set; }
        public string from { get; set; }
        public string fromString { get; set; }
        public string to { get; set; }
        public string toString { get; set; }
        public object tmpFromAccountId { get; set; }
        public string tmpToAccountId { get; set; }
    }

    public class History
    {
        public string id { get; set; }
        public Author author { get; set; }
        public DateTime created { get; set; }
        public List<Item> items { get; set; }
    }

    public class Changelog
    {
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public int total { get; set; }
        public List<History> histories { get; set; }
    }

    public class Component
    {
        public string self { get; set; }
        public string id { get; set; }
        public string name { get; set; }
    }

    public class Creator
    {
        public string self { get; set; }
        public string accountId { get; set; }
        public AvatarUrls avatarUrls { get; set; }
        public string displayName { get; set; }
        public bool active { get; set; }
        public string timeZone { get; set; }
        public string accountType { get; set; }
    }

    public class Customfield10020
    {
        public int id { get; set; }
        public string name { get; set; }
        public string state { get; set; }
        public int boardId { get; set; }
        public string goal { get; set; }
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }
    }

    public class Project
    {
        public string self { get; set; }
        public string id { get; set; }
        public string key { get; set; }
        public string name { get; set; }
        public string projectTypeKey { get; set; }
        public bool simplified { get; set; }
        public AvatarUrls avatarUrls { get; set; }
    }

    public class Reporter
    {
        public string self { get; set; }
        public string accountId { get; set; }
        public AvatarUrls avatarUrls { get; set; }
        public string displayName { get; set; }
        public bool active { get; set; }
        public string timeZone { get; set; }
        public string accountType { get; set; }
    }

    public class Priority
    {
        public string self { get; set; }
        public string iconUrl { get; set; }
        public string name { get; set; }
        public string id { get; set; }
    }

    public class Timetracking
    {
        public string remainingEstimate { get; set; }
        public string timeSpent { get; set; }
        public int remainingEstimateSeconds { get; set; }
        public int timeSpentSeconds { get; set; }
    }

    public class UpdateAuthor
    {
        public string self { get; set; }
        public string accountId { get; set; }
        public AvatarUrls avatarUrls { get; set; }
        public string displayName { get; set; }
        public bool active { get; set; }
        public string timeZone { get; set; }
        public string accountType { get; set; }
    }

    public class Comment2
    {
        public string self { get; set; }
        public string id { get; set; }
        public Author author { get; set; }
        public string body { get; set; }
        public UpdateAuthor updateAuthor { get; set; }
        public DateTime created { get; set; }
        public DateTime updated { get; set; }
        public bool jsdPublic { get; set; }
    }

    public class Comment
    {
        public List<Comment2> comments { get; set; }
        public string self { get; set; }
        public int maxResults { get; set; }
        public int total { get; set; }
        public int startAt { get; set; }
    }

    public class Worklog2
    {
        public string self { get; set; }
        public Author author { get; set; }
        public UpdateAuthor updateAuthor { get; set; }
        public string comment { get; set; }
        public DateTime created { get; set; }
        public DateTime updated { get; set; }
        public DateTime started { get; set; }
        public string timeSpent { get; set; }
        public int timeSpentSeconds { get; set; }
        public string id { get; set; }
        public string issueId { get; set; }
    }

    public class Worklog
    {
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public int total { get; set; }
        public List<Worklog2> worklogs { get; set; }
    }

    public class Assignee
    {
        public string self { get; set; }
        public string accountId { get; set; }
        public AvatarUrls avatarUrls { get; set; }
        public string displayName { get; set; }
        public bool active { get; set; }
        public string timeZone { get; set; }
        public string accountType { get; set; }
    }

    public class StatusCategory
    {
        public string self { get; set; }
        public int id { get; set; }
        public string key { get; set; }
        public string colorName { get; set; }
        public string name { get; set; }
    }

    public class Status
    {
        public string self { get; set; }
        public string description { get; set; }
        public string iconUrl { get; set; }
        public string name { get; set; }
        public string id { get; set; }
        public StatusCategory statusCategory { get; set; }
    }

    public class Fields
    {
        public string summary { get; set; }
        public List<Component> components { get; set; }
        public Creator creator { get; set; }
        public DateTime created { get; set; }
        public string description { get; set; }
        public List<Customfield10020> customfield_10020 { get; set; }
        public Project project { get; set; }
        public Reporter reporter { get; set; }
        public Priority priority { get; set; }
        public double customfield_10024 { get; set; }
        public List<object> labels { get; set; }
        public Timetracking timetracking { get; set; }
        public Comment comment { get; set; }
        public Worklog worklog { get; set; }
        public Assignee assignee { get; set; }
        public DateTime updated { get; set; }
        public Status status { get; set; }
    }

    public class Issue
    {
        public string expand { get; set; }
        public string id { get; set; }
        public string self { get; set; }
        public string key { get; set; }
        public Changelog changelog { get; set; }
        public Fields fields { get; set; }
    }

    public class JiraProjectResponse
    {
        public string id { get; set; }
        public string key { get; set; }
        public string name { get; set; }
        public string projectTypeKey { get; set; }
        public bool simplified { get; set; }
        public string style { get; set; }
        public bool isPrivate { get; set; }
    }

}
