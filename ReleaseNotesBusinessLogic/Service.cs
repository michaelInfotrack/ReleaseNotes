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
using System.Configuration;

namespace ReleaseNotesBusinessLogic
{
    public class Service
    {
        private Jira _jira;
        private string _releaseLabelToday;
        public string ReleaseLabelToday { get { return _releaseLabelToday; } }

        public Service(string jiraPath, string username, string password)
        {
            // Pull crednetials from Config instead (later)
            _jira = GetJiraConnection(jiraPath, username, password);
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
            catch(Exception e)
            {
                return new List<Issue>();
            }
        }

        public string CreateIssuesHistory(List<Tuple<string, List<Issue>>> listTuple)
        {
            var automation = new WordDocumentAutomation();

            string outputPath, inputPath;

            outputPath = String.Format(@"\\syd-schfile01-t\Images\ReleaseNotes\ReleaseNotes_{0}.pdf", DateTime.Today.ToString("yyyy-MM-dd"));
            inputPath = @"\\syd-schfile01-t\Images\ReleaseNotes\ReleaseNotes.docx";

            automation.AddToHistoryIssues(inputPath, outputPath, listTuple);

            return outputPath;
        }


        public string CreateIssuesHistory(string label)
        {
            var automation = new WordDocumentAutomation();
            var doclabel = label;
            DateTime date;
            bool isLabelDate = IsLabelDate(label, out date);
            string outputPath, inputPath;

            outputPath = String.Format(@"\\syd-schfile01-t\Images\ReleaseNotes\ReleaseNotes_{0}.pdf", DateTime.Today.ToString("yyyy-MM-dd"));
            inputPath = @"\\syd-schfile01-t\Images\ReleaseNotes\ReleaseNotes.docx";
            doclabel = isLabelDate ? date.ToLongDateString() : label;

            var tuple = new List<Tuple<string, List<Issue>>>();
            tuple.Add(new Tuple<string, List<Issue>>(doclabel, GetDailyReleaseIssues(label)));

            automation.AddToHistoryIssues(inputPath, outputPath, tuple); 

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
