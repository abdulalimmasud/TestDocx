using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace testDocx
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string templatePath = @"C:\Users\Alim\Desktop\Templates\test1\Irrigation-Template - Copy.docx";
                string contentPath = @"C:\Users\Alim\Desktop\Templates\test1\Irrigation-Content.docx";
                //TestInterOp.TestInterOpWord(templatePath, contentPath);

                Document tempDocument = new Document(templatePath);
                Document contentDocument = new Document(contentPath);

                //tempDocument.Sections.ToList()[0].GetText().Replace("[content]", contentDocument.GetText());

                Regex regex = new Regex("[content]", RegexOptions.IgnoreCase);
                


                FindMatchedNodes obj = new FindMatchedNodes();

                
                

                NodeCollection bodyNodes = tempDocument.GetChildNodes(NodeType.Body, true);
                Node tempBody = bodyNodes.FirstOrDefault();

                string txt = tempBody.GetText();
                //FindandInsertDocument replacedoc = new FindandInsertDocument(contentPath);
                //
                //tempDocument.Range.Replace(regex, replacedoc, false);

                //tempDocument.GetText().Replace("[content]", contentDocument.GetText());

                //var paragraphs = tempDocument.FirstSection.Range.Replace(regex, );






                //var dstNode = tempDocument.GetChildNodes(NodeType.Body, true).FirstOrDefault();


                //tempDocument.AppendDocument(contentDocument, ImportFormatMode.KeepSourceFormatting);

                //tempDocument.Save(@"C:\Users\Alim\Desktop\Templates\test1\test.docx");

                Console.Read();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
    public class FindMatchedNodes : IReplacingCallback

    {

        //Store Matched nodes in ArrayList

        public ArrayList nodes = new ArrayList();


        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)

        {

            // This is a Run node that contains either the beginning or the complete match.

            Node currentNode = e.MatchNode;


            // The first (and may be the only) run can contain text before the match,

            // in this case it is necessary to split the run.

            if (e.MatchOffset > 0)

                currentNode = SplitRun((Run)currentNode, e.MatchOffset);


            ArrayList runs = new ArrayList();


            // Find all runs that contain parts of the match string.

            int remainingLength = e.Match.Value.Length;

            while (


            (remainingLength > 0) &&


            (currentNode != null) &&


            (currentNode.GetText().Length <= remainingLength))

            {


                runs.Add(currentNode);


                remainingLength = remainingLength - currentNode.GetText().Length;


                // Select the next Run node.

                // Have to loop because there could be other nodes such as BookmarkStart etc.

                do

                {

                    currentNode = currentNode.NextSibling;

                }

                while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));

            }


            // Split the last run that contains the match if there is any text left.

            if ((currentNode != null) && (remainingLength > 0))

            {

                SplitRun((Run)currentNode, remainingLength);

                runs.Add(currentNode);

            }


            String runText = "";

            foreach (Run run in runs)

                runText += run.Text;


            ((Run)runs[0]).Text = runText;


            for (int i = 1; i < runs.Count; i++)

            {

                ((Run)runs[i]).Remove();

            }


            nodes.Add(runs[0]);


            // Signal to the replace engine to do nothing because we have already done all what we wanted.

            return ReplaceAction.Skip;

        }
        private static Run SplitRun(Run run, int position)
        {
            Run afterRun = (Run)run.Clone(true);
            afterRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);
            return afterRun;
        }
    }
    public class FindandInsertDocument : IReplacingCallback
    {
        private String path;

        public FindandInsertDocument(String documentpath)
        {
            path = documentpath;
        }

        /// <summary>
        /// This method is called by the Aspose.Words find and replace engine for each match.
        /// </summary>
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // This is a Run node that contains either the beginning or the complete match.
            Node currentNode = e.MatchNode;

            // The first (and may be the only) run can contain text before the match,
            // in this case it is necessary to split the run.
            if (e.MatchOffset > 0)
                currentNode = SplitRun((Run)currentNode, e.MatchOffset);

            // This array is used to store all nodes of the match for further removing.
            ArrayList runs = new ArrayList();

            // Find all runs that contain parts of the match string.
            int remainingLength = e.Match.Value.Length;
            while (
                (remainingLength > 0) &&
                (currentNode != null) &&
                (currentNode.GetText().Length <= remainingLength))
            {
                runs.Add(currentNode);
                remainingLength = remainingLength - currentNode.GetText().Length;

                // Select the next Run node.
                // Have to loop because there could be other nodes such as BookmarkStart etc.
                do
                {
                    currentNode = currentNode.NextSibling;
                }
                while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
            }

            // Split the last run that contains the match if there is any text left.
            if ((currentNode != null) && (remainingLength > 0))
            {
                SplitRun((Run)currentNode, remainingLength);
                runs.Add(currentNode);
            }

            // Create Document Builder and insert document
            DocumentBuilder builder = new DocumentBuilder(e.MatchNode.Document as Document);
            builder.MoveTo((Run)runs[runs.Count - 1]);


            Document doc = new Document(path);
            builder.InsertDocument(doc, ImportFormatMode.KeepSourceFormatting);

            // Now remove all runs in the sequence.
            foreach (Run run in runs)
                run.Remove();

            // Signal to the replace engine to do nothing because we have already done all what we wanted.
            return ReplaceAction.Skip;
        }

        private static Run SplitRun(Run run, int position)
        {
            Run afterRun = (Run)run.Clone(true);
            afterRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);
            return afterRun;
        }
    }
}
