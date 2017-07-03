using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Atlassian.Jira;

namespace ReleaseNotesBusinessLogic
{
    public class WordDocumentAutomation
    {

        public void AddToHistoryIssues(XElement documentData, string inputFilePath, string outputFilePath, List<Issue> issues)
        {
            //Create new word object
            Microsoft.Office.Interop.Word.Application word = WordApplicationHelpers.CreateNewWordInstance();

            // process the file
            using (WordDocument wordDocWrapper = new WordDocument(word, inputFilePath))
            {
                if (wordDocWrapper.AddTagIssues(issues, DateTime.Today.AddDays(-1).ToLongDateString()))
                {
                    wordDocWrapper.SaveAs(inputFilePath, WdSaveFormat.wdFormatDocumentDefault);                    
                }
                wordDocWrapper.SaveAs(outputFilePath, WdSaveFormat.wdFormatPDF);
            }
        }
    }      
}
