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

        public void AddToHistoryIssues(string inputFilePath, string outputFilePath, List<Issue> issues, string label, bool isLabelDate)
        {
            //Create new word object
            Microsoft.Office.Interop.Word.Application word = WordApplicationHelpers.CreateNewWordInstance();

            // process the file
            using (WordDocument wordDocWrapper = new WordDocument(word, inputFilePath))
            {
                //Only save the word document if the label is a date
                if (wordDocWrapper.AddTagIssues(issues, label, isLabelDate) && isLabelDate)
                {
                    wordDocWrapper.SaveAs(inputFilePath, WdSaveFormat.wdFormatDocumentDefault);                    
                }
                wordDocWrapper.SaveAs(outputFilePath, WdSaveFormat.wdFormatPDF);
            }
        }
    }      
}
