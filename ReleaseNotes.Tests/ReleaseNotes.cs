using System;
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
            var temp = new Service();

            temp.SetJiraConnection();

        }
    }
}
