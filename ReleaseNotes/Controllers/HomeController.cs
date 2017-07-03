using System;
using System.Web.Mvc;
using System.Web.Services.Description;
using ReleaseNotes.Models;
using Outlook = Microsoft.Office.Interop.Outlook;
using Service = ReleaseNotesBusinessLogic.Service;
using System.Linq;
using System.Runtime.InteropServices;

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
            var model = new ResultsModel { JiraIssues = results.ToList() };
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
            var model = new ResultsModel { JiraIssues = result.ToList() };

            return View("Index", model);
        }

        private string GetEmailBody(ResultsModel model)
        {
            try
            {
                string body = string.Empty;
                var newLine = "<br />";
                body += "Releases for: " + DateTime.Today.ToShortDateString() + newLine + newLine;

                bool dmtHeadingAdded = false;
                bool globalHeadingAdded = false;
                bool webHeadingAdded = false;
                bool pexaHeadingAdded = false;
                bool pencilHeadingAdded = false;
                bool otherHeadingAdded = false;

                //Add the release notes here
                foreach (var item in model.JiraIssues)
                {
                    switch (item.Project)
                    {
                        case "DMT":
                            if (!dmtHeadingAdded)
                            {
                                body += AddHeading("Project: Development Management Team");
                                dmtHeadingAdded = true;
                            }

                            body += item.Key.Value + " - " + item.Summary + newLine;

                            break;

                        case "GLOB":
                            if (!globalHeadingAdded)
                            {
                                body += AddHeading("Project: Global Platform");
                                globalHeadingAdded = true;
                            }

                            body += item.Key.Value + " - " + item.Summary + newLine;

                            break;

                        case "VOI":
                            if (!globalHeadingAdded)
                            {
                                body += AddHeading("Project: VOI");
                                globalHeadingAdded = true;
                            }

                            body += item.Key.Value + " - " + item.Summary + newLine;

                            break;

                        case "WEB":
                            if (!webHeadingAdded)
                            {
                                body += AddHeading("Project: Web");
                                webHeadingAdded = true;
                            }

                            body += item.Key.Value + " - " + item.Summary + newLine;

                            break;

                        case "PEXA":
                            if (!pexaHeadingAdded)
                            {
                                body += AddHeading("Project: Pexa");
                                pexaHeadingAdded = true;
                            }

                            body += item.Key.Value + " - " + item.Summary + newLine;

                            break;

                        case "PEN":
                            if (!pencilHeadingAdded)
                            {
                                body += AddHeading("Project: Pencil");
                                pencilHeadingAdded = true;
                            }

                            body += item.Key.Value + " - " + item.Summary + newLine;

                            break;
                        default:
                            if (!otherHeadingAdded)
                            {
                                body += AddHeading("Project: Other");
                                otherHeadingAdded = true;
                            }
                            body += item.Key.Value + " - " + item.Summary + newLine;

                            break;
                    }
                }

                return body;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        private static string AddHeading(string heading)
        {
            var newLine = "<br />";
            return "<b>" + heading + "</b>" + newLine;
        }

        [HttpPost]
        public ActionResult GenerateEmail()
        {
            var result = _service.GetDailyReleaseIssues();
            var model = new ResultsModel { JiraIssues = result.ToList() };

            Outlook.Application _objApp;
            Outlook.MailItem _objMail = null;

            try
            {
                if (ModelState.IsValid)
                {

                    _objApp = new Outlook.Application();

                    _objMail = (Outlook.MailItem)_objApp.CreateItem(Outlook.OlItemType.olMailItem);
                    _objMail.To = "test@infotrack.com.au"; //Replace with InfotrackDevelopmentNotifications@infotrack.com.au from appSettings
                    _objMail.Subject = "Release Notes - ";

                    _objMail.HTMLBody = GetEmailBody(model);
                    _objMail.Display(true);
                }
            }
            catch (Exception e)
            {
                ModelState.AddModelError("Email", "An error occurred trying to open the email client.");

            }

            _objMail?.Close(Outlook.OlInspectorClose.olDiscard);

            return View("Index", model);
        }
    }
}