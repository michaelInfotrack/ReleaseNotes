using System;
using System.Web.Mvc;
using System.Web.Services.Description;
using ReleaseNotes.Models;
using Outlook = Microsoft.Office.Interop.Outlook;
using Service = ReleaseNotesBusinessLogic.Service;
using System.Linq;
using System.Runtime.InteropServices;
using Atlassian.Jira;
using ReleaseNotesBusinessLogic;

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

        private string GetEmailBody(ResultsModel model, string releaseLabel)
        {

            try
            {
                #region Variables
                string body = string.Empty;
                var newLine = "<br />";
                body += "Releases for: " + releaseLabel + newLine + newLine; //This should probably be the releaseLabel

                bool dmtHeadingAdded = false;
                bool globalHeadingAdded = false;
                bool voiHeadingAdded = false;
                bool iMajorHeadingAdded = false;
                bool infotrackUkHeadingAdded = false;
                bool internalHeadingAdded = false;
                bool labsHeadingAdded = false;
                bool mapItHeadingAdded = false;
                bool mapleHeadingAdded = false;
                bool pencilHeadingAdded = false;
                bool pexaHeadingAdded = false;
                bool planItHeadingAdded = false;
                bool revealHeadingAdded = false;
                bool settleItHeadingAdded = false;
                bool signItHeadingAdded = false;
                bool testHeadingAdded = false;
                bool trackItHeadingAdded = false;
                bool usListHeadingAdded = false;
                bool usPlatformHeadingAdded = false;
                bool webHeadingAdded = false;
                bool weCareHeadingAdded = false;
                bool otherHeadingAdded = false;
                #endregion

                //Add the release notes here
                foreach (var item in model.JiraIssues)
                {
                    switch (item.Project)
                    {
                        #region Projects
                        case "DMT":
                            if (!dmtHeadingAdded)
                            {
                                body += AddHeading("Project: Development Management Team");
                                dmtHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "GLOB":
                            if (!globalHeadingAdded)
                            {
                                body += AddHeading("Project: Global Platform");
                                globalHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "VOI":
                            if (!voiHeadingAdded)
                            {
                                body += AddHeading("Project: VOI");
                                voiHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "IMAJOR":
                            if (!iMajorHeadingAdded)
                            {
                                body += AddHeading("Project: iMajor");
                                iMajorHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "UK":
                            if (!infotrackUkHeadingAdded)
                            {
                                body += AddHeading("Project: Infotrack UK");
                                infotrackUkHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "IN":
                            if (!internalHeadingAdded)
                            {
                                body += AddHeading("Project: Internal");
                                internalHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "LABS":
                            if (!labsHeadingAdded)
                            {
                                body += AddHeading("Project: LABS");
                                labsHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "MAPIT":
                            if (!mapItHeadingAdded)
                            {
                                body += AddHeading("Project: MapIT");
                                mapItHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "MAP":
                            if (!mapleHeadingAdded)
                            {
                                body += AddHeading("Project: Maple");
                                mapleHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "PEN":
                            if (!pencilHeadingAdded)
                            {
                                body += AddHeading("Project: Pencil");
                                pencilHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "PEXA":
                            if (!pexaHeadingAdded)
                            {
                                body += AddHeading("Project: Pexa");
                                pexaHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "PLN":
                            if (!planItHeadingAdded)
                            {
                                body += AddHeading("Project: PlanIT");
                                planItHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "REV":
                            if (!revealHeadingAdded)
                            {
                                body += AddHeading("Project: Reveal");
                                revealHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "SC":
                            if (!settleItHeadingAdded)
                            {
                                body += AddHeading("Project: SettleIT");
                                settleItHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "SIG":
                            if (!signItHeadingAdded)
                            {
                                body += AddHeading("Project: SignIT");
                                signItHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "TEST":
                            if (!testHeadingAdded)
                            {
                                body += AddHeading("Project: Test");
                                testHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "TIT":
                            if (!trackItHeadingAdded)
                            {
                                body += AddHeading("Project: TrackIT");
                                trackItHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "UL":
                            if (!usListHeadingAdded)
                            {
                                body += AddHeading("Project: US - The List");
                                usListHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "USP":
                            if (!usPlatformHeadingAdded)
                            {
                                body += AddHeading("Project: US Platform");
                                usPlatformHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "WEB":
                            if (!webHeadingAdded)
                            {
                                body += AddHeading("Project: Website");
                                webHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        case "CAR":
                            if (!weCareHeadingAdded)
                            {
                                body += AddHeading("Project: We Care");
                                weCareHeadingAdded = true;
                            }

                            body = AddItemToBody(body, item);

                            break;

                        default:
                            if (!otherHeadingAdded)
                            {
                                body += AddHeading("Project: Other");
                                otherHeadingAdded = true;
                            }
                            body = AddItemToBody(body, item);

                            break;
#endregion
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

        private static string AddItemToBody(string body, Issue item)
        {
            var newLine = "<br />";
            var url = "https://infotrack.atlassian.net/browse/";
            body += string.Format("<a href='{0}'>"+ item.Key.Value + "</a>", url + item.Key.Value) + " - " + item.Summary + newLine;
            return body;
        }

        private static string AddHeading(string heading)
        {
            var newLine = "<br />";
            return "<b>" + heading + "</b>" + newLine;
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

                    DateTime date;
                    string dateString = releaseLabel;
                    if (_service.IsLabelDate(releaseLabel, out date))
                    {
                        dateString = date.ToString("dd/MM/yyyy");
                    }


                    System.Diagnostics.Process[] outlookProcess = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");

                    //Check if an existing outlook process already exsist, if so, use it.
                    _objApp = outlookProcess.Length != 0
                        ? Marshal.GetActiveObject("Outlook.Application") as Outlook.Application
                        : new Outlook.Application();

                    _objMail = (Outlook.MailItem)_objApp.CreateItem(Outlook.OlItemType.olMailItem);
                    _objMail.To = "test@infotrack.com.au"; //Replace with InfotrackDevelopmentNotifications@infotrack.com.au from appSettings
                    _objMail.Attachments.Add(_service.CreateIssuesHistory(releaseLabel));
                    _objMail.Subject = "Release Notes - ";

                    _objMail.HTMLBody = GetEmailBody(model);
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