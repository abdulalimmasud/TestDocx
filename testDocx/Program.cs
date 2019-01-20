using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace testDocx
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string templatePath = @"C:\Users\Alim\Desktop\Templates\test1\164 - Copy.docx";
                string contentPath = @"C:\Users\Alim\Desktop\Templates\test1\165.docx";

                using (var stream = System.IO.File.Open(templatePath, FileMode.Open))
                using (var doc = WordprocessingDocument.Open(stream, true))
                {
                    var mdp = doc.MainDocumentPart;
                    var id = "AltChunkId1";
                    var conPara = mdp.Document.Body.Elements<W.Paragraph>().Where(x => x.InnerText == "[content]").FirstOrDefault();

                    var conStream = System.IO.File.Open(contentPath, FileMode.Open);
                    
                    var chunk = mdp.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, id);
                    chunk.FeedData(conStream);
                    

                    var alterChunk = new AltChunk { Id = id };
                    conPara.InsertBeforeSelf(alterChunk);
                    conPara.Remove();
                    conPara.RemoveAllChildren();
                    doc.Save();

                    //XmlDocument xml = new XmlDocument();
                    //string _byteOrderMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
                    //foreach (var item in mdp.HeaderParts)
                    //{
                    //    try
                    //    {
                    //        xml.LoadXml(item.Header.InnerXml);
                    //    }
                    //    catch (Exception e)
                    //    {

                    //    }
                    //}
                    
                    //try
                    //{
                    //    xml.LoadXml(mdp.Document.InnerXml);
                    //}
                    //catch (Exception e)
                    //{

                    //}
                    //foreach (var item in mdp.FooterParts)
                    //{
                    //    try
                    //    {
                    //        xml.LoadXml(item.Footer.InnerXml);
                    //    }
                    //    catch (Exception e)
                    //    {

                    //    }
                    //}

                }

                //ValidateWordDocument(templatePath);
                //
                //ValidateCorruptedWordDocument(contentPath);

                Console.Read();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void ValidateCorruptedWordDocument(string filepath)
        {
            // Insert some text into the body, this would cause Schema Error
            using (WordprocessingDocument wordprocessingDocument =
            WordprocessingDocument.Open(filepath, true))
            {
                // Insert some text into the body, this would cause Schema Error
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
                Run run = new Run(new Text("some text"));
                body.Append(run);

                try
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    int count = 0;
                    foreach (ValidationErrorInfo error in
                        validator.Validate(wordprocessingDocument))
                    {
                        count++;
                        Console.WriteLine("Error " + count);
                        Console.WriteLine("Description: " + error.Description);
                        Console.WriteLine("ErrorType: " + error.ErrorType);
                        Console.WriteLine("Node: " + error.Node);
                        Console.WriteLine("Path: " + error.Path.XPath);
                        Console.WriteLine("Part: " + error.Part.Uri);
                        Console.WriteLine("-------------------------------------------");
                    }

                    Console.WriteLine("count={0}", count);
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
        public static void ValidateWordDocument(string filepath)
        {
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
            {
                try
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    int count = 0;
                    foreach (ValidationErrorInfo error in validator.Validate(wordprocessingDocument))
                    {
                        count++;
                        Console.WriteLine("Error " + count);
                        Console.WriteLine("Description: " + error.Description);
                        Console.WriteLine("ErrorType: " + error.ErrorType);
                        Console.WriteLine("Node: " + error.Node);
                        Console.WriteLine("Path: " + error.Path.XPath);
                        Console.WriteLine("Part: " + error.Part.Uri);
                        Console.WriteLine("-------------------------------------------");
                    }

                    Console.WriteLine("count={0}", count);
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                wordprocessingDocument.Close();
            }
        }
    }
    //public class FindMatchedNodes : IReplacingCallback

    //{

    //    //Store Matched nodes in ArrayList

    //    public ArrayList nodes = new ArrayList();


    //    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)

    //    {

    //        // This is a Run node that contains either the beginning or the complete match.

    //        Node currentNode = e.MatchNode;


    //        // The first (and may be the only) run can contain text before the match,

    //        // in this case it is necessary to split the run.

    //        if (e.MatchOffset > 0)

    //            currentNode = SplitRun((Run)currentNode, e.MatchOffset);


    //        ArrayList runs = new ArrayList();


    //        // Find all runs that contain parts of the match string.

    //        int remainingLength = e.Match.Value.Length;

    //        while (


    //        (remainingLength > 0) &&


    //        (currentNode != null) &&


    //        (currentNode.GetText().Length <= remainingLength))

    //        {


    //            runs.Add(currentNode);


    //            remainingLength = remainingLength - currentNode.GetText().Length;


    //            // Select the next Run node.

    //            // Have to loop because there could be other nodes such as BookmarkStart etc.

    //            do

    //            {

    //                currentNode = currentNode.NextSibling;

    //            }

    //            while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));

    //        }


    //        // Split the last run that contains the match if there is any text left.

    //        if ((currentNode != null) && (remainingLength > 0))

    //        {

    //            SplitRun((Run)currentNode, remainingLength);

    //            runs.Add(currentNode);

    //        }


    //        String runText = "";

    //        foreach (Run run in runs)

    //            runText += run.Text;


    //        ((Run)runs[0]).Text = runText;


    //        for (int i = 1; i < runs.Count; i++)

    //        {

    //            ((Run)runs[i]).Remove();

    //        }


    //        nodes.Add(runs[0]);


    //        // Signal to the replace engine to do nothing because we have already done all what we wanted.

    //        return ReplaceAction.Skip;

    //    }
    //    private static Run SplitRun(Run run, int position)
    //    {
    //        Run afterRun = (Run)run.Clone(true);
    //        afterRun.Text = run.Text.Substring(position);
    //        run.Text = run.Text.Substring(0, position);
    //        run.ParentNode.InsertAfter(afterRun, run);
    //        return afterRun;
    //    }
    //}
    //public class FindandInsertDocument : IReplacingCallback
    //{
    //    private String path;

    //    public FindandInsertDocument(String documentpath)
    //    {
    //        path = documentpath;
    //    }

    //    /// <summary>
    //    /// This method is called by the Aspose.Words find and replace engine for each match.
    //    /// </summary>
    //    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
    //    {
    //        // This is a Run node that contains either the beginning or the complete match.
    //        Node currentNode = e.MatchNode;

    //        // The first (and may be the only) run can contain text before the match,
    //        // in this case it is necessary to split the run.
    //        if (e.MatchOffset > 0)
    //            currentNode = SplitRun((Run)currentNode, e.MatchOffset);

    //        // This array is used to store all nodes of the match for further removing.
    //        ArrayList runs = new ArrayList();

    //        // Find all runs that contain parts of the match string.
    //        int remainingLength = e.Match.Value.Length;
    //        while (
    //            (remainingLength > 0) &&
    //            (currentNode != null) &&
    //            (currentNode.GetText().Length <= remainingLength))
    //        {
    //            runs.Add(currentNode);
    //            remainingLength = remainingLength - currentNode.GetText().Length;

    //            // Select the next Run node.
    //            // Have to loop because there could be other nodes such as BookmarkStart etc.
    //            do
    //            {
    //                currentNode = currentNode.NextSibling;
    //            }
    //            while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
    //        }

    //        // Split the last run that contains the match if there is any text left.
    //        if ((currentNode != null) && (remainingLength > 0))
    //        {
    //            SplitRun((Run)currentNode, remainingLength);
    //            runs.Add(currentNode);
    //        }

    //        // Create Document Builder and insert document
    //        DocumentBuilder builder = new DocumentBuilder(e.MatchNode.Document as Document);
    //        builder.MoveTo((Run)runs[runs.Count - 1]);


    //        Document doc = new Document(path);
    //        builder.InsertDocument(doc, ImportFormatMode.KeepSourceFormatting);

    //        // Now remove all runs in the sequence.
    //        foreach (Run run in runs)
    //            run.Remove();

    //        // Signal to the replace engine to do nothing because we have already done all what we wanted.
    //        return ReplaceAction.Skip;
    //    }

    //    private static Run SplitRun(Run run, int position)
    //    {
    //        Run afterRun = (Run)run.Clone(true);
    //        afterRun.Text = run.Text.Substring(position);
    //        run.Text = run.Text.Substring(0, position);
    //        run.ParentNode.InsertAfter(afterRun, run);
    //        return afterRun;
    //    }
    //}
}
