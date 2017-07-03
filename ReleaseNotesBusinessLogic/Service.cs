using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Atlassian;
using Atlassian.Jira;
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

            CreateIssuesHistory();
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

            return _jira.Issues.GetIsssuesFromJqlAsync(jqlQuery, 100, 0, new System.Threading.CancellationToken()).Result.ToList();
        }


        private void CreateIssuesHistory()
        {
            var automation = new WordDocumentAutomation();           

            automation.AddToHistoryIssues(null, @"D:\Git\ReleaseNotes\ReleaseNotes.docx", @"D:\Git\ReleaseNotes\test.pdf", GetDailyReleaseIssues());

        }


        public object GetFormattedReleaseLabelFromDate()
        {
            throw new NotImplementedException();
        }
    }
}
