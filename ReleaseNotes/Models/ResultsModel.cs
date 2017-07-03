using Atlassian.Jira;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReleaseNotes.Models
{
    public class ResultsModel
    {
        public List<Issue> JiraIssues { get; set; }
    }
}