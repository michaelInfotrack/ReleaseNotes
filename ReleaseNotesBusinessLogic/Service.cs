using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Atlassian;
using Atlassian.Jira;
using Atlassian.Jira.Linq;
using Newtonsoft.Json;
using System.Globalization;

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

        public string GetDailyReleaseLabel()
        {
            return DateTime.Today.ToString("yyyMMdd");
        }

        public List<Issue> GetDailyReleaseIssues(string releaseLabel = "")
        {
            try
            {
                var label = string.IsNullOrEmpty(releaseLabel) ? _releaseLabelToday : releaseLabel;
                var jqlQuery = string.Format("labels = {0} ", label);

                return ExecuteJqlQuery(jqlQuery).OrderBy(i => i.Key.ToString()).ToList();
            }
            catch(Exception)
            {
                return new List<Issue>();
            }
        }


        public string CreateIssuesHistory(string label)
        {
            var automation = new WordDocumentAutomation();
            var doclabel = label;
            DateTime date;
            bool isLabelDate = IsLabelDate(label, out date);

            string outputPath = String.Format(@"\\syd-schfile01-t\Images\ReleaseNotes\ReleaseNotes_{0}.pdf", DateTime.Today.ToString("yyyy-MM-dd"));
            var inputPath = @"\\syd-schfile01-t\Images\ReleaseNotes\ReleaseNotes.docx";

            if (!isLabelDate)
            {
                outputPath = String.Format(@"\\syd-schfile01-t\Images\ReleaseNotes\ReleaseNotes_{0}_{1}.pdf", label, DateTime.Today.ToString("yyyy-MM-dd"));
                inputPath = @"\\syd-schfile01-t\Images\ReleaseNotes\ReleaseNotes_empty.docx";
            }
            else doclabel = date.ToLongDateString();

            automation.AddToHistoryIssues(inputPath, outputPath, GetDailyReleaseIssues(label), doclabel, isLabelDate);

            return outputPath;
        }

        public bool IsLabelDate(string label, out DateTime date)
        {
            return DateTime.TryParseExact(label,
                       "yyyyMMdd",
                       CultureInfo.InvariantCulture,
                       DateTimeStyles.None,
                       out date);
        }


        private List<Issue> ExecuteJqlQuery(string jqlQuery)
        {

            return _jira.Issues.GetIsssuesFromJqlAsync(jqlQuery, 100, 0, new System.Threading.CancellationToken()).Result.ToList();
        }


    }
}
