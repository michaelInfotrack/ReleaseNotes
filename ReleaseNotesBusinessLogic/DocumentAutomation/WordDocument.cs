using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Atlassian.Jira;


namespace ReleaseNotesBusinessLogic
{
    public class WordDocument:IDisposable
    {
        private Application _word;
        private Microsoft.Office.Interop.Word.Document _wordDoc;
        private string _inputFilePath;
        private string _inputFolderPath;

        private static string tempfolderLocation;
        private static object oMissing = System.Reflection.Missing.Value;


        public WordDocument(Application word, string givenInputFilePath, bool useTempFolder = false)
        {
            tempfolderLocation = Path.GetTempPath();

            this._inputFolderPath = System.IO.Path.Combine(tempfolderLocation, Guid.NewGuid().ToString());
            System.IO.Directory.CreateDirectory(this._inputFolderPath);

            this._inputFilePath = System.IO.Path.Combine(this._inputFolderPath, "Tempfile.doc");
            
            //copy the file to temp folder
            System.IO.File.Copy(givenInputFilePath, this._inputFilePath);

            this._wordDoc = word.Documents.Open(this._inputFilePath, true);
            this._word = word;

           if (_wordDoc.ProtectionType != Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
                _wordDoc.Unprotect();
        }

        public void AddTagIssues(List<Issue> list, string releaseDate)
        {           
            if (list.Count > 0)
            {
                _wordDoc.Paragraphs.Add(_wordDoc.Range(0, 0));

                var pDate = _wordDoc.Paragraphs.Add(_wordDoc.Paragraphs[2].Range);
                pDate.Format.SpaceAfter = 10f;
                pDate.Range.Text = String.Format(releaseDate);
                pDate.Range.Font.Size = 14;
                pDate.Range.Font.Name = "Arial";

                pDate.Range.InsertParagraphAfter();

                var pTable = _wordDoc.Paragraphs.Add(_wordDoc.Paragraphs[4].Range);
                pTable.Format.SpaceAfter = 10f;
                
                var table = _wordDoc.Tables.Add(pTable.Range, list.Count + 1, 4, ref oMissing, ref oMissing);

                table.Spacing = 0.5f;
                table.Columns[1].SetWidth(70, WdRulerStyle.wdAdjustNone);
                table.Columns[2].SetWidth(300,WdRulerStyle.wdAdjustNone);
                table.Columns[3].SetWidth(100, WdRulerStyle.wdAdjustNone);

                foreach (var item in list)
                {
                    for (int i = 1; i < 4; i++)
                    {
                        table.Rows[list.IndexOf(item) + 1].Cells[i].Range.Font.Size = 10;
                        table.Rows[list.IndexOf(item) + 1].Cells[i].Range.Font.Name = "Arial";
                    }

                    object address = @"https://infotrack.atlassian.net/browse/" + item.Key;
                    _wordDoc.Hyperlinks.Add(table.Rows[list.IndexOf(item) + 1].Cells[1].Range, ref address, ref oMissing, ref oMissing, item.Key.Value, ref oMissing);
                    table.Rows[list.IndexOf(item) + 1].Cells[2].Range.Text = item.Summary;                                     
                    table.Rows[list.IndexOf(item) + 1].Cells[3].Range.Text = item.Assignee;
                }
            }
        }       


        public void SaveAs(string outputFilePath, WdSaveFormat saveFormat)
        {
            
            // save the output file
            object oMissing = System.Reflection.Missing.Value;
            object oFalse = new object();
            object oFilename = outputFilePath;
            object fileFormat = saveFormat;
            _wordDoc.SaveAs( ref oFilename,
                            ref fileFormat, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            int x = _word.BackgroundSavingStatus;
               
        }

        private bool disposed = false;

        public virtual void Dispose(bool disposing)
        {          
            
            object oMissing = System.Reflection.Missing.Value;

            if (!this.disposed)
            {
                if (disposing)
                {
                    if (this._wordDoc != null)
                    {
                        this._wordDoc.Close(false, ref oMissing, ref oMissing);
                        this._wordDoc = null;
                    }

                    if (this._word != null)
                    {
                        this._word.Quit(false, ref oMissing, ref oMissing);
                        this._word = null;
                    }
                }
            }
            this.disposed = true;

            //Delete copied input file
            System.IO.File.Delete(this._inputFilePath);
            //Delete folder
            System.IO.Directory.Delete(this._inputFolderPath);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
