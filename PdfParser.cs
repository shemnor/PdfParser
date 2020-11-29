using System;
using System.Collections.Generic;
using PdfGeom = iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser.Filter;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Annot;
using iText.Layout;
using iText.Layout.Element;
using iText.Kernel.Pdf.Annot.DA;
using iText.Kernel.Colors;
using System.Drawing;
using System.Text.RegularExpressions;
using PdfColors = iText.Kernel.Colors;
using iText.Kernel.Pdf.Colorspace;
using System.Text;
using System.Diagnostics;
using System.Globalization;
using SysColor = System.Drawing.Color;




namespace PdfParser
{
    public static class PdfParser
    {
        #region ###### General methods
        private static PdfDocument getPdfForReadAndWrite(string sourcePath, string destPath)
        {
            PdfReader reader = new PdfReader(sourcePath);
            PdfWriter writer = new PdfWriter(destPath);

            PdfDocument pdfDocForReadAndWrite = new PdfDocument(reader, writer);
            return pdfDocForReadAndWrite;
        }
        private static PdfDocument getPdfForRead(string path)
        {
            PdfReader reader = new PdfReader(path);

            PdfDocument pdfDocForRead = new PdfDocument(reader);
            return pdfDocForRead;
        }
        private static PdfDocument getPdfForWrite(string path)
        {
            PdfWriter writer = new PdfWriter(path);

            PdfDocument pdfDocForWrite = new PdfDocument(writer);
            return pdfDocForWrite;
        }
        #endregion

        #region ###### copy annotations by color
        public static void copyAnnotsByColor(string sourcePath, string destPath, string tempDestPath)
        {
            //get the drawing to copy from (soruce drawing)
            PdfDocument sourceDoc = getPdfForRead(sourcePath);

            //get the new drawing for read and temporary fro write (destination drawing)
            PdfDocument destDoc = getPdfForReadAndWrite(destPath,tempDestPath);

            //get annotations from source
            PdfPage sourcePage = sourceDoc.GetFirstPage();
            IList<PdfAnnotation> annotations = sourcePage.GetAnnotations();

            //filter annotations by color
            SysColor filterColor = SysColor.Red;
            filterAnnotationsByColor(ref annotations,filterColor);

            //add annotations in destination drawing
            PdfPage destPage = destDoc.GetFirstPage();
            foreach(PdfAnnotation annot in annotations)
            {
                PdfObject annotObject = annot.GetPdfObject().CopyTo(destDoc);
                PdfAnnotation newAnnot = PdfAnnotation.MakeAnnotation(annotObject);
                destPage.AddAnnotation(newAnnot);
            }

            //close documents
            sourceDoc.Close();
            destDoc.Close();
        }
        private static void filterAnnotationsByColor(ref IList<PdfAnnotation> annots, SysColor filterColor)
        {
            for (int i = annots.Count - 1; i >= 0; i--)
            {
                PdfAnnotation annot = annots[i] as PdfAnnotation;
                SysColor annotColor = getAnnotationColors(annot);
                if (annotColor!= SysColor.White)
                {
                    if (!checkIfColorIsSimilar(filterColor, annotColor, 20))
                    {
                        annots.RemoveAt(i);
                    }
                }

            }
        }
        private static void excludeAnnotationsByColor(ref IList<PdfAnnotation> annots, SysColor filterColor)
        {
            for (int i = annots.Count - 1; i >= 0; i--)
            {
                PdfAnnotation annot = annots[i] as PdfAnnotation;
                SysColor annotColor = getAnnotationColors(annot);
                if (annotColor != null)
                {
                    if (checkIfColorIsSimilar(filterColor, annotColor, 20))
                    {
                        annots.RemoveAt(i);
                    }
                }

            }
        }
        private static SysColor getAnnotationColors(PdfAnnotation annot)
        {
            SysColor colorNull = SysColor.White;
            if (annot.GetSubtype() == PdfName.FreeText)
            {
                PdfMarkupAnnotation markupAnnot = annot as PdfMarkupAnnotation;
                string richText = markupAnnot.GetRichText().ToString();

                //THIS CHECK IS POOR WHEN CERTAINTY IS NECCESARY IE COPYING ONLY RED ANNOTS
                //FORCING RED COLOR ON ANNOTS WITH ANY RED FORMAT IS CONSERVATIVE WHEN BLACKENING ANNOTS

                //A MORE COMPREHENSIVE CHECK IS REQUIRED FOR THE "COPY RED ANNOTS" FUNCTION

                //check if Annot contains any RED - 
                if (richText.Contains("#FF"))
                {
                    return SysColor.Red;
                }

                //check if annot contains any black
                if (richText.Contains("#000000"))
                {
                    return SysColor.Black;
                }

            }
            else if (annot.GetSubtype() == PdfName.PolyLine || annot.GetSubtype() == PdfName.Line || annot.GetSubtype() == PdfName.Polygon)
            {
                double[] tempRGB = annot.GetColorObject().ToDoubleArray();
                int[] RGB = new int[3];
                for (int j = 0; j < RGB.Length; j++)
                {
                    RGB[j] = System.Convert.ToInt16(tempRGB[j] * 255);
                }
                SysColor color = SysColor.FromArgb(RGB[0], RGB[1], RGB[2]);
                return color;
            }
            return colorNull;
        }
        private static bool checkIfColorIsSimilar(SysColor baseColorRgb, SysColor newColorRgb, int tolerance)
        {
            //RGB is a 3d space. each color is a point in the 255 domain.
            //calculate distance from the base color and check against tolerance

            int distance = ((newColorRgb.R - baseColorRgb.R) ^ 2 + (newColorRgb.G - baseColorRgb.G) ^ 2 + (newColorRgb.B - baseColorRgb.B) ^ 2) ^ (1 / 2);
            return (Math.Abs(distance) <= tolerance);
        }
        #endregion

        #region ####### update content
        public static void updateDrawingWithFcrAndRevision (string sourcePath, string tempPath, string fcrNumber, string revision)
        {
            //get pdf for read and write
            PdfDocument pdfDoc = getPdfForReadAndWrite(sourcePath, tempPath);
            
            //get page
            PdfPage page = pdfDoc.GetFirstPage();
            
            //get annotation on page
            IList<PdfAnnotation> annotations = page.GetAnnotations();

            //update revision boxes
            string revRegexPattern = @"[A-Z]\.[0-9][0-9]\s*";
            updateTextboxContentWithRegex(ref annotations, revision, revRegexPattern);

            //update FCR texboxes
            string fcrRegexPattern = @"[Ff][Cc][Rr]\s[0-9][0-9][0-9][0-9][0-9][0-9]\s*";
            updateTextboxContentWithRegex(ref annotations, fcrNumber,fcrRegexPattern);

            pdfDoc.Close();
        }
        public static void changeColorOfAnnotsByColor(string sourcePath, string tempPath, System.Drawing.Color existingColor, System.Drawing.Color newColor)
        {
            //get pdf for read and write
            PdfDocument pdfDoc = getPdfForReadAndWrite(sourcePath, tempPath);

            //get page
            PdfPage page = pdfDoc.GetFirstPage();

            //get annotation on page
            IList<PdfAnnotation> annotations = page.GetAnnotations();

            //modify TextBox color
            updateAnnotsWithColor(ref annotations, existingColor, newColor);

            pdfDoc.Close();
        }
        private static void updateAnnotsWithColor(ref IList<PdfAnnotation> annotations, SysColor existingColor, SysColor newColor)
        {

            //filter annotations by color
            filterAnnotationsByColor(ref annotations, existingColor);

            for (int i = 0; i < annotations.Count; i++)
            {
                PdfAnnotation annot = annotations[i] as PdfAnnotation;
                if (annot.GetSubtype() == PdfName.FreeText)
                {
                    modifyTextboxColor(ref annot, newColor);
                }
                else if (annot.GetSubtype() == PdfName.PolyLine || annot.GetSubtype() == PdfName.Line || annot.GetSubtype() == PdfName.Polygon)
                {
                    modifyGeomAnnotColor(ref annot, newColor);
                }
            }
        }
        private static void updateTextboxContentWithRegex(ref IList<PdfAnnotation> annotations, string newContent, string regexPattern)
        {
            for (int i = 0; i < annotations.Count; i++)
            {
                if (annotations[i].GetSubtype() == PdfName.FreeText)
                {
                    //get content
                    PdfAnnotation annot = annotations[i] as PdfAnnotation;
                    string contents = annot.GetContents().ToString();

                    if (Regex.IsMatch(contents, regexPattern))
                    {
                        //get color 
                        SysColor color = SysColor.Red;

                        //modify contents
                        modifyTextboxContents(ref annot, newContent, color);
                    }
                }
            }
        }
        private static void modifyTextboxContents(ref PdfAnnotation annot, string contents, SysColor defaultColor)
        {
            //get color
            string textColorHex = String.Format("#{0}{1}{2}", defaultColor.R.ToString("X2"), defaultColor.B.ToString("X2"), defaultColor.G.ToString("X2"));

            //get font size
            string contentRT = annot.GetPdfObject().Get(PdfName.RC).ToString();
            int fontSizeIndex = contentRT.LastIndexOf("font-size:") + 10;
            int fontSize = int.Parse(contentRT.Substring(fontSizeIndex, 2));

            //get richText
            string richText = generateRichText(fontSize, textColorHex, contents.TrimEnd(' '));

            //get modification date
            string time = String.Format("D:{0}'00'", DateTime.Now.ToString("yyyyMMddHHmmsszz", DateTimeFormatInfo.InvariantInfo));

            //get rotation
            //PdfNumber rotation = annot.GetPdfObject().Get(PdfName.Rotate) as PdfNumber;

            //modify free text annot
            modifyTextBox(ref annot, contents, defaultColor, fontSize, richText, time);
        }
        private static void modifyTextboxColor(ref PdfAnnotation annot, SysColor defaultColor)
        {
            //get contents
            string contents = annot.GetContents().ToString();

            //get color
            string textColorHex = String.Format("#{0}{1}{2}", defaultColor.R.ToString("X2"), defaultColor.B.ToString("X2"), defaultColor.G.ToString("X2"));

            //get font size
            string contentRT = annot.GetPdfObject().Get(PdfName.RC).ToString();
            int fontSizeIndex = contentRT.LastIndexOf("font-size:") + 10;
            int fontSize = int.Parse(contentRT.Substring(fontSizeIndex, 2));

            //get richText
            string richText = generateRichText(fontSize, textColorHex, contents.TrimEnd(' '));

            //get modification date
            string time = String.Format("D:{0}'00'", DateTime.Now.ToString("yyyyMMddHHmmsszz", DateTimeFormatInfo.InvariantInfo));

            //get rotation
            //PdfNumber rotation = annot.GetPdfObject().Get(PdfName.Rotate) as PdfNumber;

            //modify free text annot
            modifyTextBox(ref annot, contents, defaultColor, fontSize, richText, time);

        }
        private static void modifyTextBox(ref PdfAnnotation annotation, string contents, SysColor defaultColor, int fontSize, string richText, string modificationDate)
        {
            PdfDictionary annotDict = annotation.GetPdfObject();
            PdfFreeTextAnnotation annotFT = annotation as PdfFreeTextAnnotation;
            PdfMarkupAnnotation annotMA = annotation as PdfMarkupAnnotation;

            //remove the AP
            annotDict.Remove(PdfName.AP);

            //set contents
            annotation.SetContents(contents);

            //*set default appearance
            AnnotationDefaultAppearance DA = new AnnotationDefaultAppearance();
            DA.SetColor(new DeviceRgb(defaultColor));
            DA.SetFont(StandardAnnotationFont.HelveticaBold);
            DA.SetFontSize(fontSize);
            annotFT.SetDefaultAppearance(DA);

            //*set default Style
            string contentDS = String.Format("font: Helvetica ,sans - serif {0}.00pt; color:{1}", fontSize, ColorTranslator.ToHtml(defaultColor));
            annotDict.Put(PdfName.DS, new PdfString(contentDS));

            //*set modification date
            annotation.SetDate(new PdfString(modificationDate));

            //*set Rich text
            annotMA.SetRichText(new PdfString(richText));

        }
        private static string generateRichText(int fontSize, string textColorHex, string content)
        {
            //create richtext
            string richText = String.Format(
                "<?xml version=\"1.0\"?>" +
                "<body xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:xfa=\"http://www.xfa.org/schema/xfa-data/1.0/\" xfa:APIVersion=\"Acrobat:11.0.0\" xfa:spec=\"2.0.2\">" +
                "<p dir=\"ltr\">" +
                "<span style=\"text-align:left;font-size:{0}pt;font-style:normal;font-weight:bold;color:{1};font-family:Helvetica\">{2}" +
                "</span>" +
                "</p>" +
                "</body>", fontSize, textColorHex, content);

            return richText;
        }
        private static void modifyGeomAnnotColor(ref PdfAnnotation annot, SysColor newColor)
        {
            float[] floatRGB = new float[3];
            for (int j = 0; j < 2; j++)
            {
                floatRGB[j] = (float)System.Convert.ToDouble(newColor.R) / 255;
            }
            annot.SetColor(floatRGB);
        }
        #endregion

        #region create new content
        private static PdfAnnotation createNewTextbox(PdfArray rectangle, string contents, SysColor defaultColor, string richText, string author, string creationDate, PdfNumber rotation, int fontSize)
        {
            PdfAnnotation annotation = new PdfFreeTextAnnotation(new PdfGeom.Rectangle(1, 1, 1, 1), null);
            PdfFreeTextAnnotation annotFT = annotation as PdfFreeTextAnnotation;
            PdfMarkupAnnotation annotMA = annotation as PdfMarkupAnnotation;
            PdfDictionary annotDict = annotation.GetPdfObject();

            //set border effect
            PdfDictionary BE = new PdfDictionary();
            PdfDictionary contentsBE = new PdfDictionary();
            contentsBE.Put(PdfName.S, PdfName.S);
            BE.Put(PdfName.BE, contentsBE);
            annotDict.Put(PdfName.BE, BE);

            //set border style
            PdfDictionary BS = new PdfDictionary();
            PdfDictionary contentsBS = new PdfDictionary();
            contentsBS.Put(PdfName.S, PdfName.S);
            contentsBS.Put(PdfName.W, new PdfNumber(1));
            BS.Put(PdfName.BS, contentsBS);
            annotDict.Put(PdfName.BS, BS);

            //set background color
            annotation.SetColor(ColorConstants.WHITE);

            //set the opacity
            annotDict.Put(PdfName.CA, new PdfNumber(1));

            //set content
            annotation.SetContents(contents);

            //set creation date 
            annotDict.Put(PdfName.CreationDate, new PdfString(creationDate));

            //*set default appearance
            AnnotationDefaultAppearance DA = new AnnotationDefaultAppearance();
            DA.SetColor(new DeviceRgb(defaultColor));
            DA.SetFont(StandardAnnotationFont.HelveticaBold);
            DA.SetFontSize(fontSize);
            annotFT.SetDefaultAppearance(DA);

            //*set default Style
            string contentDS = String.Format("font: Helvetica ,sans - serif {0}.00pt; color:{1}", fontSize, ColorTranslator.ToHtml(defaultColor));
            annotDict.Put(PdfName.DS, new PdfString(contentDS));

            //set flag
            annotation.SetFlag(4);

            //*set modification date
            annotation.SetDate(new PdfString(creationDate));

            //set quadding
            annotDict.Put(PdfName.Q, new PdfNumber(0));

            //*set Rich text
            annotMA.SetRichText(new PdfString(richText));

            //set rectangle
            annotation.SetRectangle(rectangle);

            //set Rotation
            annotDict.Put(PdfName.Rotate, rotation);

            //set subject
            annotDict.Put(PdfName.Subj, new PdfString("Textbox"));

            //set title
            annotation.SetTitle(new PdfString(author));

            //set type
            annotDict.Put(PdfName.Type, PdfName.Annot);

            return annotation;
        }
        #endregion

        #region analise content
        private static void getTheNewestVevision()
        {

        }
        #endregion

        #region ################################################## WORK IN PROGRESS


        private static string pdfText(string path)
        {
            PdfReader reader = new PdfReader(path);
            PdfDocument pdfDoc = new PdfDocument(reader);
            string text = string.Empty;
            for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
            {
                //text += PdfTextExtractor.GetTextFromPage(reader, page);
            }
            reader.Close();
            return text;
        }
        private static string pdfTextByRect(string path)
        {
            PdfReader reader = new PdfReader(path);
            PdfDocument pdfDoc = new PdfDocument(reader);
            iText.Kernel.Geom.Rectangle rect = new iText.Kernel.Geom.Rectangle(2871,326,440,90);
            TextRegionEventFilter regionFilter = new TextRegionEventFilter(rect);
            for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
            {
                ITextExtractionStrategy strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), regionFilter);
                try
                {
                    String str = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy) + "\n";
                    return str;
                    //fos.write(str.GetBytes("UTF-8"));
                }
                catch (Exception e)
                {
                    return null;
                }
            }
            return null;
        }
        private static void testfromInternet(string sourcePath, string destPath, string tempDestPath)
        {
            //get the drawing to copy from (soruce drawing)
            PdfDocument sourceDoc = getPdfForRead(sourcePath);

            //get the new drawing for read and temporary fro write (destination drawing)
            PdfDocument masterPdfDoc = getPdfForReadAndWrite(destPath, tempDestPath);
            PdfPage masterDocPage = masterPdfDoc.GetFirstPage();

            //get annotations from source
            PdfPage sourcePage = sourceDoc.GetFirstPage();
            IList<PdfAnnotation> annotations = sourcePage.GetAnnotations();

            foreach (var anno in annotations)
            {
                if(anno.GetSubtype() == PdfName.FreeText)
                {
                    if (anno.GetContents().ToString().Equals("CJ"))
                    {
                        //get color 
                        System.Drawing.Color color = System.Drawing.Color.Red;
                        string textColor = String.Format("#{0}{1}{2}", color.R.ToString("X2"), color.B.ToString("X2"), color.G.ToString("X2"));
                        
                        //get content
                        string content = "FCR 111111";

                        //get rectangle
                        PdfArray oldRect = anno.GetRectangle();

                        //get font size
                        string contentRT = anno.GetPdfObject().Get(PdfName.RC).ToString();
                        int fontSizeIndex = contentRT.LastIndexOf("font-size:") + 10;
                        int fontSize = int.Parse(contentRT.Substring(fontSizeIndex, 2));

                        //create richtext
                        string richText = String.Format(
                            "<?xml version=\"1.0\"?>" +
                            "<body xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:xfa=\"http://www.xfa.org/schema/xfa-data/1.0/\" xfa:APIVersion=\"Acrobat:11.0.0\" xfa:spec=\"2.0.2\">" +
                            "<p dir=\"ltr\">" +
                            "<span style=\"text-align:left;font-size:{0}pt;font-style:normal;font-weight:bold;color:{1};font-family:Helvetica\">{2}" +
                            "</span>" +
                            "</p>" +
                            "</body>", fontSize, textColor, content);

                        //get username
                        string user = Environment.UserName;

                        //get time
                        string time = String.Format("D:{0}'00'", DateTime.Now.ToString("yyyyMMddHHmmsszz", DateTimeFormatInfo.InvariantInfo));

                        //get rotation
                        PdfNumber rotation = anno.GetPdfObject().Get(PdfName.Rotate) as PdfNumber;



                        //create the annotation
                        PdfAnnotation pdfAnnotation = createNewTextbox(oldRect, content, color, richText, user, time, rotation, fontSize);

                        //add annotation
                        masterDocPage.AddAnnotation(pdfAnnotation);

                        //delete old annotation
                        masterDocPage.RemoveAnnotation(anno);
                    }
                }

            }

            sourceDoc.Close();
            masterPdfDoc.Close();
        }
        private static String ExtractAnnotationText(PdfStream xObject)
        {
            PdfResources resources = new PdfResources(xObject.GetAsDictionary(PdfName.Resources));
            ITextExtractionStrategy strategy = new LocationTextExtractionStrategy();

            PdfCanvasProcessor processor = new PdfCanvasProcessor(strategy);
            processor.ProcessContent(xObject.GetBytes(), resources);
            var text = strategy.GetResultantText();
            return text;
        }
        private static void addText(string path)
        {
            PdfWriter writer = new PdfWriter(path);
            PdfDocument pdfDocument = new PdfDocument(writer);
            PdfPage page = pdfDocument.GetFirstPage();

            PdfGeom.Rectangle rectangle = new PdfGeom.Rectangle(20, 100, 100, 100);
            PdfAnnotation newAnnot = new PdfFreeTextAnnotation(rectangle, new PdfString("TEST"));
            newAnnot.SetColor(new float[] { 0, 0, 0 });
            page.AddAnnotation(newAnnot);

            pdfDocument.Close();
        }
        private static int getAnotationsCount(string path)
        {
            PdfReader reader = new PdfReader(path);
            PdfDocument pdfDoc = new PdfDocument(reader);
            Document doc = new Document(pdfDoc);
            PdfPage page = pdfDoc.GetPage(1);
            IList<PdfAnnotation> annotations = page.GetAnnotations();
            return annotations.Count;
        }
        #endregion
    }
}
