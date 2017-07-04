using System;
using System.Configuration;
using System.Web.Mvc;
using System.Web.Services.Description;
using ReleaseNotes.Models;
using Outlook = Microsoft.Office.Interop.Outlook;
using Service = ReleaseNotesBusinessLogic.Service;
using System.Linq;
using System.Runtime.InteropServices;
using Atlassian.Jira;
using ReleaseNotesBusinessLogic;
using static ReleaseNotes.Models.ResultsModel;
using System.Reflection;
using System.ComponentModel;

namespace ReleaseNotes.Controllers
{
    public class HomeController : Controller
    {
        private Service _service;

        public HomeController()
        {
            _service = new Service();
        }

        public ActionResult Index()
        {
            var results = _service.GetDailyReleaseIssues();
            var model = new ResultsModel { JiraIssues = results };
            return View(model);
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpGet]
        public ActionResult GetByDate(string releaseLabel)
        {
            var result = _service.GetDailyReleaseIssues(releaseLabel);
            var model = new ResultsModel { JiraIssues = result };

            return View("Index", model);
        }

        private static string GetEnumDescription(Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());

            DescriptionAttribute[] attributes =
                (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0)
                return attributes[0].Description;
            else
                return value.ToString();
        }

        private string GetEmailBody(ResultsModel model, string releaseLabel)
        {
            try
            {
                #region Variables
                string body = string.Empty;
                var newLine = "<br />";
                DateTime releaseDate;
                var isLabelDate = _service.IsLabelDate(releaseLabel, out releaseDate);

                body += @"<!DOCTYPE HTML PUBLIC "" -//W3C//DTD HTML 4.01 Transitional//EN""><html><head><title> LEAP Disbursements Invoice</title><meta http - equiv = ""Content-Type"" content = ""text/html; charset=iso-8859-1""></head><body>";
                body += "Releases for: " + releaseDate.ToLongDateString() + newLine + newLine; //This should probably be the releaseLabel

                #endregion

                foreach (var projectIssues in model.JiraIssues.GroupBy(x => x.Project).ToList())
                {
                    body += @"<table width=""500px;"" border=""0"" cellspacing=""2"" cellpadding=""2"">";
                    body += @"<tr width=""100px;"">";
                    body += String.Format(@"<th align=""left"" colspan=""2"" >{0}</th>", GetEnumDescription((ProjectTypes)System.Enum.Parse(typeof(ProjectTypes), projectIssues.FirstOrDefault().Project)));
                    body += "</tr>";
                    body += "</tr></tr>";
                    foreach (var issue in projectIssues.ToList())
                    {
                        body += "<tr>";
                        body = AddItemToBody(body, issue);
                        body += "</tr>";
                    }
                    body += "</table>";
                    body += "<br />";
                } 
                

                return body;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        private static string AddItemToBody(string body, Issue item)
        {
            var url = "https://infotrack.atlassian.net/browse/";
            body += string.Format(@"<td><a href='{0}'>" + item.Key.Value + "</a></td>", url + item.Key.Value) + @"<td align=""left"" >" + item.Summary  + @"</td>";
            return body;
        }


        [HttpPost]
        public ActionResult GenerateEmail(string releaseLabel)
        {
            if (releaseLabel == string.Empty)
            {
                releaseLabel = _service.GetDailyReleaseLabel();
            }

            var result = _service.GetDailyReleaseIssues(releaseLabel);
            var model = new ResultsModel { JiraIssues = result };

            Outlook.Application _objApp;
            Outlook.MailItem _objMail = null;

            try
            {
                if (ModelState.IsValid)
                {
                    System.Diagnostics.Process[] outlookProcess = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");

                    //Check if an existing outlook process already exsist, if so, use it.
                    _objApp = outlookProcess.Length != 0
                        ? Marshal.GetActiveObject("Outlook.Application") as Outlook.Application
                        : new Outlook.Application();

                    _objMail = (Outlook.MailItem)_objApp.CreateItem(Outlook.OlItemType.olMailItem);
                    _objMail.To = ConfigurationManager.AppSettings["SendToEmail"];
                    _objMail.Attachments.Add(_service.CreateIssuesHistory(releaseLabel));
                    _objMail.Subject = "Release Notes";

                    _objMail.HTMLBody = GetEmailBody(model, releaseLabel);
                    _objMail.Display(true);
                }
            }
            catch (Exception e)
            {
                ViewBag.Message = e.Message;
                _objMail?.Close(Outlook.OlInspectorClose.olDiscard);
            }

            return View("Index", model);
        }
    }
}