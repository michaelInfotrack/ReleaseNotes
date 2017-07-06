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

        public void AddToHistoryIssues(string inputFilePath, string outputFilePath, List<Tuple<string, List<Issue>>> listTuple)
        {
            //Create new word object
            Microsoft.Office.Interop.Word.Application word = WordApplicationHelpers.CreateNewWordInstance();

            // process the file
            using (WordDocument wordDocWrapper = new WordDocument(word, inputFilePath))
            {
                foreach (var item in listTuple)
                {
                    wordDocWrapper.AddTagIssues(item.Item2, item.Item1, true);
                }
                wordDocWrapper.SaveAs(outputFilePath, WdSaveFormat.wdFormatPDF);
            }
        }
    }      
}
