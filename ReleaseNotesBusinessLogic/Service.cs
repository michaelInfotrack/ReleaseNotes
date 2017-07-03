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

        string sampleRequestJson = @"{
    'name': 'All Open Bugs',
    'description': 'Lists all open bugs',
    'jql': 'type = Bug and resolution is empty',
    'favourite': true,
    'favouritedCount': 0}";
        private Jira _jiraConnection;


        public Service()
        {
            SetJiraConnection();
        }

        public void SetJiraConnection()
        {
            var jira = Jira.CreateRestClient("https://infotrack.atlassian.net", "michael.lachlan", "Password1");

            var issues = from i in jira.Issues.Queryable
                         orderby i.Created
                         select i;


           

        }




        public string RunQueryRequest()
        {


            return null;


        }

    }
}
