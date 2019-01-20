using System;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace testDocx
{
    //var contentDoc = app.Documents.Open(contentPath);
    //contentDoc.Activate();
    //contentDoc.ActiveWindow.Selection.WholeStory();
    //contentDoc.ActiveWindow.Selection.Copy();
    //contentDoc.Close();
    //.Select();
    public static class TestInterOp
    {
        public static void TestInterOpWord(string templatePath, string contentPath)
        {
            var app = new Word.Application();
            var tempDoc = app.Documents.Open(templatePath);
            tempDoc.Activate();
            tempDoc.Content.Find.ClearFormatting();

            int index = tempDoc.Content.Text.LastIndexOf("[content]");
            int lastIndex = index + 9;
            var rng = tempDoc.Range(index, lastIndex);
            rng.Select();
            rng.Cut();
            rng.InsertFile(contentPath);
            tempDoc.Close();
            app.Quit();
        }
        public static void MailMergeTest()
        {
            var app = new Word.Application();
            var doc = app.ActiveDocument;
            
        }
    }
    public class TestMerge
    {
        //private void CombineDocuments()
        //{
        //    object wdPageBreak = 7;
        //    object wdStory = 6;
        //    object oMissing = System.Reflection.Missing.Value;
        //    object oFalse = false;
        //    object oTrue = true;
        //    string fileDirectory = @"C:documents";
        //    Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
        //    Microsoft.Office.Interop.Word.Document wDoc = WordApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        //    string[] wordFiles = Directory.GetFiles(fileDirectory, "*.doc");
        //    for (int i = 0; i < wordFiles.Length; i++)
        //    {
        //        string file = wordFiles[i];
        //        wDoc.Application.Selection.Range.InsertFile(file, ref oMissing, ref oMissing, ref oMissing, ref oFalse);
        //        wDoc.Application.Selection.Range.InsertBreak(ref wdPageBreak);
        //        wDoc.Application.Selection.EndKey(ref wdStory, ref oMissing);
        //    }
        //    string combineDocName = Path.Combine(fileDirectory, "Merged Document.doc");
        //    if (File.Exists(combineDocName))
        //        File.Delete(combineDocName);
        //    object combineDocNameObj = combineDocName;
        //    wDoc.SaveAs(ref combineDocNameObj, ref m_WordDocumentType, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        //}
        Word.Application wrdApp;
        Word._Document wrdDoc;
        Object oMissing = System.Reflection.Missing.Value;
        Object oFalse = false;
        private void InsertLines(int LineNum)
        {
            int iCount;

            // Insert "LineNum" blank lines.
            for (iCount = 1; iCount <= LineNum; iCount++)
            {
                wrdApp.Selection.TypeParagraph();
            }
        }

        private void FillRow(Word._Document oDoc, int Row, string Text1,
        string Text2, string Text3, string Text4)
        {
            // Insert the data into the specific cell.
            oDoc.Tables[1].Cell(Row, 1).Range.InsertAfter(Text1);
            oDoc.Tables[1].Cell(Row, 2).Range.InsertAfter(Text2);
            oDoc.Tables[1].Cell(Row, 3).Range.InsertAfter(Text3);
            oDoc.Tables[1].Cell(Row, 4).Range.InsertAfter(Text4);
        }

        private void CreateMailMergeDataFile()
        {
            Word._Document oDataDoc;
            int iCount;

            Object oName = "C:\\DataDoc.doc";
            Object oHeader = "FirstName, LastName, Address, CityStateZip";
            wrdDoc.MailMerge.CreateDataSource(ref oName, ref oMissing,
            ref oMissing, ref oHeader, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing);

            // Open the file to insert data.
            oDataDoc = wrdApp.Documents.Open(ref oName, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            ref oMissing/*, ref oMissing */);

            for (iCount = 1; iCount <= 2; iCount++)
            {
                oDataDoc.Tables[1].Rows.Add(ref oMissing);
            }
            // Fill in the data.
            FillRow(oDataDoc, 2, "Steve", "DeBroux",
            "4567 Main Street", "Buffalo, NY  98052");
            FillRow(oDataDoc, 3, "Jan", "Miksovsky",
            "1234 5th Street", "Charlotte, NC  98765");
            FillRow(oDataDoc, 4, "Brian", "Valentine",
            "12348 78th Street  Apt. 214",
            "Lubbock, TX  25874");
            // Save and close the file.
            oDataDoc.Save();
            oDataDoc.Close(ref oFalse, ref oMissing, ref oMissing);
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            Word.Selection wrdSelection;
            Word.MailMerge wrdMailMerge;
            Word.MailMergeFields wrdMergeFields;
            Word.Table wrdTable;
            string StrToAdd;

            // Create an instance of Word  and make it visible.
            wrdApp = new Word.Application();
            wrdApp.Visible = true;

            // Add a new document.
            wrdDoc = wrdApp.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            wrdDoc.Select();

            wrdSelection = wrdApp.Selection;
            wrdMailMerge = wrdDoc.MailMerge;

            // Create a MailMerge Data file.
            CreateMailMergeDataFile();

            // Create a string and insert it into the document.
            StrToAdd = "State University\r\nElectrical Engineering Department";
            wrdSelection.ParagraphFormat.Alignment =
            Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wrdSelection.TypeText(StrToAdd);

            InsertLines(4);

            // Insert merge data.
            wrdSelection.ParagraphFormat.Alignment =
            Word.WdParagraphAlignment.wdAlignParagraphLeft;
            wrdMergeFields = wrdMailMerge.Fields;
            wrdMergeFields.Add(wrdSelection.Range, "FirstName");
            wrdSelection.TypeText(" ");
            wrdMergeFields.Add(wrdSelection.Range, "LastName");
            wrdSelection.TypeParagraph();

            wrdMergeFields.Add(wrdSelection.Range, "Address");
            wrdSelection.TypeParagraph();
            wrdMergeFields.Add(wrdSelection.Range, "CityStateZip");

            InsertLines(2);

            // Right justify the line and insert a date field
            // with the current date.
            wrdSelection.ParagraphFormat.Alignment =
            Word.WdParagraphAlignment.wdAlignParagraphRight;

            Object objDate = "dddd, MMMM dd, yyyy";
            wrdSelection.InsertDateTime(ref objDate, ref oFalse, ref oMissing,
            ref oMissing, ref oMissing);

            InsertLines(2);

            // Justify the rest of the document.
            wrdSelection.ParagraphFormat.Alignment =
            Word.WdParagraphAlignment.wdAlignParagraphJustify;

            wrdSelection.TypeText("Dear ");
            wrdMergeFields.Add(wrdSelection.Range, "FirstName");
            wrdSelection.TypeText(",");
            InsertLines(2);

            // Create a string and insert it into the document.
            StrToAdd = "Thank you for your recent request for next " +
            "semester's class schedule for the Electrical " +
            "Engineering Department. Enclosed with this " +
            "letter is a booklet containing all the classes " +
            "offered next semester at State University.  " +
            "Several new classes will be offered in the " +
            "Electrical Engineering Department next semester.  " +
            "These classes are listed below.";
            wrdSelection.TypeText(StrToAdd);

            InsertLines(2);

            // Insert a new table with 9 rows and 4 columns.
            wrdTable = wrdDoc.Tables.Add(wrdSelection.Range, 9, 4,
            ref oMissing, ref oMissing);
            // Set the column widths.
            wrdTable.Columns[1].SetWidth(51, Word.WdRulerStyle.wdAdjustNone);
            wrdTable.Columns[2].SetWidth(170, Word.WdRulerStyle.wdAdjustNone);
            wrdTable.Columns[3].SetWidth(100, Word.WdRulerStyle.wdAdjustNone);
            wrdTable.Columns[4].SetWidth(111, Word.WdRulerStyle.wdAdjustNone);
            // Set the shading on the first row to light gray.
            wrdTable.Rows[1].Cells.Shading.BackgroundPatternColorIndex =
            Word.WdColorIndex.wdGray25;
            // Bold the first row.
            wrdTable.Rows[1].Range.Bold = 1;
            // Center the text in Cell (1,1).
            wrdTable.Cell(1, 1).Range.Paragraphs.Alignment =
            Word.WdParagraphAlignment.wdAlignParagraphCenter;

            // Fill each row of the table with data.
            FillRow(wrdDoc, 1, "Class Number", "Class Name",
            "Class Time", "Instructor");
            FillRow(wrdDoc, 2, "EE220", "Introduction to Electronics II",
            "1:00-2:00 M,W,F", "Dr. Jensen");
            FillRow(wrdDoc, 3, "EE230", "Electromagnetic Field Theory I",
            "10:00-11:30 T,T", "Dr. Crump");
            FillRow(wrdDoc, 4, "EE300", "Feedback Control Systems",
            "9:00-10:00 M,W,F", "Dr. Murdy");
            FillRow(wrdDoc, 5, "EE325", "Advanced Digital Design",
            "9:00-10:30 T,T", "Dr. Alley");
            FillRow(wrdDoc, 6, "EE350", "Advanced Communication Systems",
            "9:00-10:30 T,T", "Dr. Taylor");
            FillRow(wrdDoc, 7, "EE400", "Advanced Microwave Theory",
            "1:00-2:30 T,T", "Dr. Lee");
            FillRow(wrdDoc, 8, "EE450", "Plasma Theory",
            "1:00-2:00 M,W,F", "Dr. Davis");
            FillRow(wrdDoc, 9, "EE500", "Principles of VLSI Design",
            "3:00-4:00 M,W,F", "Dr. Ellison");

            // Go to the end of the document.
            Object oConst1 = Word.WdGoToItem.wdGoToLine;
            Object oConst2 = Word.WdGoToDirection.wdGoToLast;
            wrdApp.Selection.GoTo(ref oConst1, ref oConst2, ref oMissing, ref oMissing);
            InsertLines(2);

            // Create a string and insert it into the document.
            StrToAdd = "For additional information regarding the " +
            "Department of Electrical Engineering, " +
            "you can visit our Web site at ";
            wrdSelection.TypeText(StrToAdd);
            // Insert a hyperlink to the Web page.
            Object oAddress = "http://www.ee.stateu.tld";
            Object oRange = wrdSelection.Range;
            wrdSelection.Hyperlinks.Add(oRange, ref oAddress, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing);
            // Create a string and insert it into the document
            StrToAdd = ".  Thank you for your interest in the classes " +
            "offered in the Department of Electrical " +
            "Engineering.  If you have any other questions, " +
            "please feel free to give us a call at " +
            "555-1212.\r\n\r\n" +
            "Sincerely,\r\n\r\n" +
            "Kathryn M. Hinsch\r\n" +
            "Department of Electrical Engineering \r\n";
            wrdSelection.TypeText(StrToAdd);

            // Perform mail merge.
            wrdMailMerge.Destination = Word.WdMailMergeDestination.wdSendToNewDocument;
            wrdMailMerge.Execute(ref oFalse);

            // Close the original form document.
            wrdDoc.Saved = true;
            wrdDoc.Close(ref oFalse, ref oMissing, ref oMissing);


            // Release References.
            wrdSelection = null;
            wrdMailMerge = null;
            wrdMergeFields = null;
            wrdDoc = null;
            wrdApp = null;

            // Clean up temp file.
            System.IO.File.Delete("C:\\DataDoc.doc");
        }
    }
}
