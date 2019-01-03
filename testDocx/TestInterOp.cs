using System;
using Word = Microsoft.Office.Interop.Word;

namespace testDocx
{
    public static class TestInterOp
    {
        public static void TestInterOpWord(string templatePath, string contentPath)
        {
            var app = new Word.Application();
            var contentDoc = app.Documents.Open(contentPath);
            contentDoc.Activate();
            contentDoc.ActiveWindow.Selection.WholeStory();
            contentDoc.ActiveWindow.Selection.Copy();
            contentDoc.Close();
            var tempDoc = app.Documents.Open(templatePath);
            tempDoc.Activate();
            tempDoc.Content.Find.ClearFormatting();

            int index = tempDoc.Content.Text.LastIndexOf("[content]");
            int lastIndex = index + 9;
            tempDoc.Range(index, lastIndex).Select();
            tempDoc.ActiveWindow.Selection.Paste();
            tempDoc.Close();
        }
    }
}
