using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace testDocx
{
    public class MergeMultiple
    {
        private string wordDocumentPath;
        /// <summary>
        /// Contains key value pairs for the Word document's merge fields. The keys are the field's name and the values are the data you want to insert, it should contain a key-value pair for the FileNames
        /// </summary>
        private List<Dictionary<string, string>> MergeFields;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="WordDocumentPath">Path to the Word Document</param>
        /// <param name="MergeFields"></param>
        public MergeMultiple(string WordDocumentPath, List<Dictionary<string, string>> MergeFields)
        {
            wordDocumentPath = WordDocumentPath;
            this.MergeFields = MergeFields;
        }

        /// <summary>
        /// Performes the merge and saves the files with the given Format
        /// </summary>
        /// <param name="DestinationFolder"></param>
        /// <param name=""></param>
        public void Merge(string DestinationFolder, Microsoft.Office.Interop.Word.WdSaveFormat Format, string FileNameKey, bool CreateFolder = true)
        {
            Microsoft.Office.Interop.Word.Application wordApplication = null;
            Microsoft.Office.Interop.Word.Document wordDocument = null;

            bool FileNameKeyExists = false;
            foreach (Dictionary<string, string> FieldValuePair in MergeFields)
            {
                if (FieldValuePair.ContainsKey(FileNameKey))
                {
                    FileNameKeyExists = true;
                    break;
                }
            }
            if (!FileNameKeyExists) throw new ArgumentException("The given key for FileName doesn't exist");

            if (!Directory.Exists(DestinationFolder))
            {
                if (!CreateFolder) throw new IOException("Destination folder doesn't exist");
                else Directory.CreateDirectory(DestinationFolder);
            }
            try
            {
                wordApplication = new Microsoft.Office.Interop.Word.Application();
                foreach (Dictionary<string, string> FieldValuePair in MergeFields)
                {
                    Microsoft.Office.Interop.Word.Document MergeDocument = wordApplication.Documents.Add(wordDocumentPath);
                    Microsoft.Office.Interop.Word.Fields DocumentFields = MergeDocument.Fields;
                    //Search through fields and replace any Mergefield found
                    foreach (Microsoft.Office.Interop.Word.Field Field in DocumentFields)
                    {
                        string FieldText = Field.Code.Text;
                        if (FieldText.StartsWith(" MERGEFIELD"))
                        {
                            string FieldName = FieldText.Substring(11, FieldText.Length - 11).Trim();
                            foreach (KeyValuePair<string, string> Entry in FieldValuePair)
                            {
                                if (Entry.Key.Equals(FieldName, StringComparison.CurrentCultureIgnoreCase))
                                {
                                    Field.Select();
                                    wordApplication.Selection.TypeText(Entry.Value);
                                }
                            }
                        }
                    }
                    Marshal.ReleaseComObject(DocumentFields);

                    Microsoft.Office.Interop.Word.Sections DocumentSections = MergeDocument.Sections;
                    //Search through the headers and footers for Mergefields and replace it
                    foreach (Microsoft.Office.Interop.Word.Section Section in DocumentSections)
                    {
                        Microsoft.Office.Interop.Word.Fields HeaderFields = Section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Fields;
                        foreach (Microsoft.Office.Interop.Word.Field Field in HeaderFields)
                        {
                            string FieldText = Field.Code.Text;
                            if (FieldText.StartsWith(" MERGEFIELD"))
                            {
                                string FieldName = FieldText.Substring(11, FieldText.Length - 11).Trim();
                                foreach (KeyValuePair<string, string> Entry in FieldValuePair)
                                {
                                    if (Entry.Key.Equals(FieldName, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        Field.Select();
                                        wordApplication.Selection.TypeText(Entry.Value);
                                    }
                                }
                            }
                        }
                        Marshal.ReleaseComObject(HeaderFields);
                    }
                    Marshal.ReleaseComObject(DocumentSections);
                    MergeDocument.SaveAs2(Path.Combine(DestinationFolder, FieldValuePair[FileNameKey]), Format);
                    MergeDocument.Close(false);
                    Marshal.ReleaseComObject(MergeDocument);
                }

            }
            ///TODO
            catch (Exception)
            {
                throw;
            }
            finally
            {
                wordApplication.Quit(false);
                Marshal.ReleaseComObject(wordApplication);
            }
        }

    }
}
