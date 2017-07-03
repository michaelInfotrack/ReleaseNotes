using System;
using System.Web.Mvc;
using ReleaseNotes.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ReleaseNotes.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
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

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Contact(EmailFormModel model)
        {
            Outlook.Application _objApp;
            Outlook.MailItem _objMail;

            if (ModelState.IsValid)
            {
                _objApp = new Outlook.Application();
                _objMail = (Outlook.MailItem)_objApp.CreateItem(Outlook.OlItemType.olMailItem);
                _objMail.To = "test@infotrack.com.au"; //Replace with InfotrackDevelopmentNotifications@infotrack.com.au
                _objMail.Subject = "Release Notes - " + model.ReleaseTitle;

                _objMail.Body = GetEmailBody();
                _objMail.Display(true);
            }
            return View(model);
        }

        private string GetEmailBody()
        {
            try
            {
                string emailBody = string.Empty;

                //Add the release notes here

                return emailBody;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}