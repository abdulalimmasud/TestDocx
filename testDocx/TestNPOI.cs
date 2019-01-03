using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace testDocx
{
    public static class TestNPOI
    {
        public static void Test()
        {
            string templatePath = @"C:\Users\Alim\Desktop\Templates\test1\Irrigation-Template - Copy.docx";
            string contentPath = @"C:\Users\Alim\Desktop\Templates\test1\Irrigation-Content.docx";
            //TestInterOp.TestInterOpWord(templatePath, contentPath);

            using (FileStream templateStream = new FileStream(templatePath, FileMode.OpenOrCreate))
            {
                XWPFDocument tempDocument = new XWPFDocument(templateStream);
                var paragraphs = tempDocument.Paragraphs.ToList();

                int index = paragraphs.FindIndex(x => x.Text == "[content]");



                using (FileStream contentStream = new FileStream(contentPath, FileMode.Open))
                {
                    XWPFDocument conDoc = new XWPFDocument(contentStream);
                    var conParagraph = conDoc.Paragraphs.ToList();

                    paragraphs.InsertRange(index, conDoc.Paragraphs);
                    paragraphs.RemoveAt(index);
                    for (int i = 0; i < paragraphs.Count; i++)
                    {
                        if (i >= index)
                        {
                            tempDocument.CreateParagraph();
                        }
                        tempDocument.SetParagraph(paragraphs[i], i);
                    }
                }

                //XWPFDocument document = new XWPFDocument();
                //XWPFParagraph paragraph = document.CreateParagraph();
                //for (int i = 0; i < paragraphs.Count; i++)
                //{
                //    document.CreateParagraph();
                //    document.SetParagraph(paragraphs[i], i);
                //}
                FileStream out1 = new FileStream(@"C:\Users\Alim\Desktop\Templates\test1\test.docx", FileMode.Create);
                tempDocument.Write(out1);
                out1.Close();

            }
        }
        public static void Test1()
        {
            string templatePath = @"C:\Users\Alim\Desktop\Templates\test1\Irrigation-Template - Copy.docx";
            string contentPath = @"C:\Users\Alim\Desktop\Templates\test1\Irrigation-Content.docx";
            //TestInterOp.TestInterOpWord(templatePath, contentPath);

            XWPFDocument srcDoc = new XWPFDocument(new FileStream(templatePath, FileMode.OpenOrCreate));

            XWPFDocument destDoc = new XWPFDocument();

            // Copy document layout.
            CopyLayout(srcDoc, destDoc);

            Stream outStream = new FileStream("Destination.docx", FileMode.Create);


            foreach (IBodyElement bodyElement in srcDoc.BodyElements)
            {
                BodyElementType elementType = bodyElement.ElementType;

                if (elementType == BodyElementType.PARAGRAPH)
                {

                    XWPFParagraph srcPr = (XWPFParagraph)bodyElement;

                    //CopyStyle(srcDoc, destDoc, srcDoc.GetStyles().GetStyle(srcPr.StyleID));

                    bool hasImage = false;

                    XWPFParagraph dstPr = destDoc.CreateParagraph();

                    // Extract image from source docx file and insert into destination docx file.

                    foreach (XWPFRun srcRun in srcPr.Runs)
                    {
                        dstPr.CreateRun();

                        if (srcRun.GetEmbeddedPictures().Count > 0)
                            hasImage = true;

                        foreach (XWPFPicture pic in srcRun.GetEmbeddedPictures())
                        {
                            byte[] img = pic.GetPictureData().Data;

                            long cx = pic.GetCTPicture().spPr.xfrm.ext.cx;
                            long cy = pic.GetCTPicture().spPr.xfrm.ext.cy;

                            try
                            {
                                // Working addPicture Code below...
                                string blipId = dstPr.Document.AddPictureData(img, pic.GetHashCode());
                                destDoc.CreatePictureCxCy(blipId, destDoc.GetNextPicNameNumber(pic.GetHashCode()), cx, cy);

                            }
                            catch (FormatException e1)
                            {
                                Console.WriteLine(e1.StackTrace);
                            }
                        }


                    }



                    if (hasImage == false)
                    {
                        int pos = destDoc.Paragraphs.Count - 1;
                        destDoc.SetParagraph(srcPr, pos);
                    }

                }
                else if (elementType == BodyElementType.TABLE)
                {

                    //XWPFTable table = (XWPFTable)bodyElement;

                    //copyStyle(srcDoc, destDoc, srcDoc.getStyles().getStyle(table.getStyleID()));

                    //destDoc.createTable();

                    //int pos = destDoc.getTables().size() - 1;

                    //destDoc.setTable(pos, table);
                }
            }


            destDoc.Write(outStream);
            outStream.Close();
        }
        private static void CopyStyle(XWPFDocument srcDoc, XWPFDocument destDoc, XWPFStyle style)
        {
            if (destDoc == null || style == null)
                return;

            if (destDoc.GetCTStyle() == null)
            {
                destDoc.CreateStyles();
            }

            List<XWPFStyle> usedStyleList = srcDoc.GetStyles().GetUsedStyleList(style);

            for (int i = 0; i < usedStyleList.Count; i++)
            {
                destDoc.GetStyles().AddStyle(usedStyleList[i]);
            }
        }

        // Copy Page Layout.
        //
        // if next error message shows up, download "ooxml-schemas-1.1.jar" file and
        // add it to classpath.
        //
        // [Error]
        // The type org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar
        // cannot be resolved.
        // It is indirectly referenced from required .class files
        //
        // This error happens because there is no CTPageMar class in
        // poi-ooxml-schemas-3.10.1-20140818.jar.
        //
        // [ref.] http://poi.apache.org/faq.html#faq-N10025
        // [ref.] http://poi.apache.org/overview.html#components
        //
        // < ooxml-schemas 1.1 download >
        // http://repo.maven.apache.org/maven2/org/apache/poi/ooxml-schemas/1.1/
        //


        private static void CopyLayout(XWPFDocument srcDoc, XWPFDocument destDoc)
        {
            CT_PageMar pgMar = srcDoc.Document.body.sectPr.pgMar;

            string bottom = pgMar.bottom;
            ulong footer = pgMar.footer;
            ulong gutter = pgMar.gutter;
            ulong header = pgMar.header;
            ulong left = pgMar.left;
            ulong right = pgMar.right;
            string top = pgMar.top;

            //CT_PageMar addNewPgMar = destDoc.Document.body.sectPr.pgMar;

            //addNewPgMar.bottom = bottom;
            //addNewPgMar.footer = footer;
            //addNewPgMar.gutter = gutter;
            //addNewPgMar.setHeader(header);
            //addNewPgMar.setLeft(left);
            //addNewPgMar.setRight(right);
            //addNewPgMar.setTop(top);

            //CT_PageSz pgSzSrc = srcDoc.Document.body.sectPr.pgSz;

            //string code = pgSzSrc.code;
            //BigInteger h = pgSzSrc.getH();
            //Enum orient = pgSzSrc.getOrient();
            //BigInteger w = pgSzSrc.getW();

            //CT_PageSz addNewPgSz = destDoc.getDocument().getBody().addNewSectPr().addNewPgSz();

            //addNewPgSz.setCode(code);
            //addNewPgSz.setH(h);
            //addNewPgSz.setOrient(orient);
            //addNewPgSz.setW(w);
        }
    }
}
