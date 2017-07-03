using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Atlassian;
using Atlassian.Jira;
using Atlassian.Jira.Linq;
using Newtonsoft.Json;

namespace ReleaseNotesBusinessLogic
{
    public class Service
    {
        private Jira _jira;
        private string _releaseLabelToday;

        public Service()
        {
            // Pull crednetials from Config instead (later)
            _jira = GetJiraConnection(@"https://infotrack.atlassian.net", "michael.lachlan@infotrack.com.au", "Password2");
            _releaseLabelToday = GetDailyReleaseLabel();
        }

        private Jira GetJiraConnection(string connectionPath, string username, string password)
        {
            return Jira.CreateRestClient(connectionPath, username, password);
        }


        private string GetDailyReleaseLabel()
        {
            return DateTime.Today.ToString("yyyMMdd");
        }


        public List<Issue> GetDailyReleaseIssues(string releaseLabel = "")
        {
            var label = string.IsNullOrEmpty(releaseLabel) ? _releaseLabelToday : releaseLabel;
            var jqlQuery = string.Format("labels = {0} ", label);

            return ExecuteJqlQuery(jqlQuery);
        }

        public void GetPreviousReleases(string yearFilter = "2017")
        {
        }

        private List<Issue> ExecuteJqlQuery(string jqlQuery)
        {

            return _jira.Issues.GetIsssuesFromJqlAsync(jqlQuery, 100, 0, new System.Threading.CancellationToken()).Result.ToList();
        }


    }
}
