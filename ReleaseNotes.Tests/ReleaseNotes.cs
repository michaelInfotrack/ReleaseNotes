using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReleaseNotesBusinessLogic;

namespace ReleaseNotes.Tests
{
    [TestClass]
    public class ReleaseNotes
    {
        [TestMethod]
        public void TestMethod1()
        {
            var service = new Service();

            var issueList = service.GetDailyReleaseIssues();

        }
    }
}
