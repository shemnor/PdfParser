using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using SysColor = System.Drawing.Color;
using iText.Kernel.Colors;
using iText.Kernel.Pdf;
using iText.Kernel.Font;
using iText.Kernel.Pdf.Annot;
using iText.Kernel.Pdf.Annot.DA;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Filter;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Layout;
using PdfGeom = iText.Kernel.Geom;

namespace PdfParser
{
    public static class PdfParser
    {
        private static class RegexPattern
        {
            public static String revision { get { return @"[A-Z]\.[0-9][0-9]\s*"; } }
            public static String fcr { get { return @"[Ff][Cc][Rr]\s*[0-9][0-9][0-9][0-9][0-9][0-9]\s*"; } }
        }

        #region ###### General methods
        private static PdfDocument _getPdfForReadAndWrite(string sourcePath, string destPath)
        {
            PdfReader reader = new PdfReader(sourcePath);
            PdfWriter writer = new PdfWriter(destPath);

            PdfDocument pdfDocForReadAndWrite = new PdfDocument(reader, writer);
            return pdfDocForReadAndWrite;
        }
        private static PdfDocument _getPdfForRead(string path)
        {
            PdfReader reader = new PdfReader(path);

            PdfDocument pdfDocForRead = new PdfDocument(reader);
            return pdfDocForRead;
        }
        private static PdfDocument _getPdfForWrite(string path)
        {
            PdfWriter writer = new PdfWriter(path);

            PdfDocument pdfDocForWrite = new PdfDocument(writer);
            return pdfDocForWrite;
        }
        #endregion

        #region ###### copy annotations by color
        public static bool copyAnnotsByColor(string sourcePath, string destPath, string tempDestPath)
        {
            PdfDocument sourceDoc = null;
            PdfDocument destDoc = null;

            try
            {
                //get the drawing to copy from (soruce drawing)
                sourceDoc = _getPdfForRead(sourcePath);

                //get the new drawing for read and temporary fro write (destination drawing)
                destDoc = _getPdfForReadAndWrite(destPath, tempDestPath);

                //get annotations from source
                PdfPage sourcePage = sourceDoc.GetFirstPage();
                IList<PdfAnnotation> annotations = sourcePage.GetAnnotations();

                //filter annotations by color
                SysColor filterColor = SysColor.Red;
                _filterAnnotationsByColor(ref annotations, filterColor);

                //add annotations in destination drawing
                PdfPage destPage = destDoc.GetFirstPage();
                foreach (PdfAnnotation annot in annotations)
                {
                    PdfObject annotObject = annot.GetPdfObject().CopyTo(destDoc);
                    PdfAnnotation newAnnot = PdfAnnotation.MakeAnnotation(annotObject);
                    destPage.AddAnnotation(newAnnot);
                }

                //close documents
                sourceDoc.Close();
                destDoc.Close();

                //confirm complete
                return true;
            }
            catch
            {
                //clean up
                if (sourceDoc != null) { sourceDoc.Close(); }
                if (destDoc != null) { destDoc.Close(); }

                //confirm incomplete
                return false;
            }


        }
        private static void _filterAnnotationsByColor(ref IList<PdfAnnotation> annots, SysColor filterColor)
        {
            for (int i = annots.Count - 1; i >= 0; i--)
            {
                PdfAnnotation annot = annots[i] as PdfAnnotation;
                SysColor annotColor = _getAnnotationColors(annot);
                
                if (annotColor!= SysColor.White)
                {
                    if (!_checkIfColorIsSimilar(filterColor, annotColor, 20))
                    {
                        annots.RemoveAt(i);
                    }
                }

            }
        }
        private static void _excludeAnnotationsByColor(ref IList<PdfAnnotation> annots, SysColor filterColor)
        {
            for (int i = annots.Count - 1; i >= 0; i--)
            {
                PdfAnnotation annot = annots[i] as PdfAnnotation;
                SysColor annotColor = _getAnnotationColors(annot);
                if (annotColor != null)
                {
                    if (_checkIfColorIsSimilar(filterColor, annotColor, 20))
                    {
                        annots.RemoveAt(i);
                    }
                }

            }
        }
        private static SysColor _getAnnotationColors(PdfAnnotation annot)
        {
            SysColor colorNull = SysColor.White;
            if (annot.GetSubtype() == PdfName.FreeText)
            {
                PdfMarkupAnnotation markupAnnot = annot as PdfMarkupAnnotation;
                //if no text return white collor
                if (markupAnnot.GetRichText() == null) { return colorNull; }

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
                //check if has color
                if (annot.GetColorObject() == null) { return colorNull; }

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
        private static bool _checkIfColorIsSimilar(SysColor baseColorRgb, SysColor newColorRgb, int tolerance)
        {
            //RGB is a 3d space. each color is a point in the 255 domain.
            //calculate distance from the base color and check against tolerance

            int distance = ((newColorRgb.R - baseColorRgb.R) ^ 2 + (newColorRgb.G - baseColorRgb.G) ^ 2 + (newColorRgb.B - baseColorRgb.B) ^ 2) ^ (1 / 2);
            return (Math.Abs(distance) <= tolerance);
        }
        #endregion

        #region ####### update content
        public static bool updateDrawingTextboxByRegex (string sourcePath, string tempPath, string newContent)
        {
            PdfDocument pdfDoc = null;
            try
            {
                //get pdf for read and write
                pdfDoc = _getPdfForReadAndWrite(sourcePath, tempPath);

                //get page
                PdfPage page = pdfDoc.GetFirstPage();

                //get annotation on page
                IList<PdfAnnotation> annotations = page.GetAnnotations();

                if (Regex.IsMatch(newContent, RegexPattern.revision)) { _updateTextboxContentWithRegex(ref annotations, newContent, RegexPattern.revision ); }
                else if(Regex.IsMatch(newContent,RegexPattern.fcr)) { _updateTextboxContentWithRegex(ref annotations, newContent, RegexPattern.fcr); }

                //close document
                pdfDoc.Close();

                //confirm complete
                return true;
            }
            catch
            {
                //clean up
                if (pdfDoc != null) { pdfDoc.Close(); }

                //confirm incomplete
                return false;
            }
        }
        public static bool changeColorOfAnnotsByColor(string sourcePath, string tempPath, SysColor existingColor, SysColor newColor)
        {
            PdfDocument pdfDoc = null;
            try
            {
                //get pdf for read and write
                pdfDoc = _getPdfForReadAndWrite(sourcePath, tempPath);

                //get page
                PdfPage page = pdfDoc.GetFirstPage();

                //get annotation on page
                IList<PdfAnnotation> annotations = page.GetAnnotations();

                //modify TextBox color
                _updateAnnotsWithColor(ref annotations, existingColor, newColor);

                //close doc
                pdfDoc.Close();

                //confirm complete
                return true;
            }
            catch
            {
                //clean up
                if (pdfDoc != null) { pdfDoc.Close(); }

                //confirm incomplete
                return false;
            }
        }
        private static void _updateAnnotsWithColor(ref IList<PdfAnnotation> annotations, SysColor existingColor, SysColor newColor)
        {

            //filter annotations by color
            _filterAnnotationsByColor(ref annotations, existingColor);

            for (int i = 0; i < annotations.Count; i++)
            {
                PdfAnnotation annot = annotations[i] as PdfAnnotation;
                if (annot.GetSubtype() == PdfName.FreeText)
                {
                    _modifyTextboxColor(ref annot, newColor);
                }
                else if (annot.GetSubtype() == PdfName.PolyLine || annot.GetSubtype() == PdfName.Line || annot.GetSubtype() == PdfName.Polygon)
                {
                    _modifyGeomAnnotColor(ref annot, newColor);
                }
            }
        }
        private static void _updateTextboxContentWithRegex(ref IList<PdfAnnotation> annotations, string newContent, string regexPattern)
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
                        _modifyTextboxContents(ref annot, newContent, color);
                    }
                }
            }
        }
        private static void _modifyTextboxContents(ref PdfAnnotation annot, string contents, SysColor defaultColor)
        {
            //get font size
            string contentRT = annot.GetPdfObject().Get(PdfName.RC).ToString();
            int fontSizeIndex = contentRT.LastIndexOf("font-size:") + 10;
            int fontSize = int.Parse(contentRT.Substring(fontSizeIndex, 2));

            //get richText
            string richText = _generateRichText(fontSize, defaultColor, contents.TrimEnd(' '));

            //get modification date
            string time = String.Format("D:{0}'00'", DateTime.Now.ToString("yyyyMMddHHmmsszz", DateTimeFormatInfo.InvariantInfo));;

            //modify free text annot
            _modifyTextBox(ref annot, contents, defaultColor, fontSize, richText, time);
        }
        private static void _modifyTextboxColor(ref PdfAnnotation annot, SysColor defaultColor)
        {
            //get contents
            string contents = annot.GetContents().ToString();

            //get font size
            string contentRT = annot.GetPdfObject().Get(PdfName.RC).ToString();
            int fontSizeIndex = contentRT.LastIndexOf("font-size:") + 10;
            int fontSize = int.Parse(contentRT.Substring(fontSizeIndex, 2));

            //get richText
            string richText = _generateRichText(fontSize, defaultColor, contents.TrimEnd(' '));

            //get modification date
            string time = String.Format("D:{0}'00'", DateTime.Now.ToString("yyyyMMddHHmmsszz", DateTimeFormatInfo.InvariantInfo));

            //modify free text annot
            if (annot.GetPdfObject().Get(PdfName.Subj).ToString().Equals("Textbox"))
            {
                _modifyTextBox(ref annot, contents, defaultColor, fontSize, richText, time);
            }
            else if (annot.GetPdfObject().Get(PdfName.Subj).ToString().Equals("Typewriter"))
            {
                _modifyTypewriter(ref annot, contents, defaultColor, fontSize, richText, time);
            }

        }
        private static void _modifyTextBox(ref PdfAnnotation annotation, string contents, SysColor defaultColor, int fontSize, string richText, string modificationDate)
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
        private static void _modifyTypewriter(ref PdfAnnotation annotation, string contents, SysColor defaultColor, int fontSize, string richText, string modificationDate)
        {
            PdfDictionary annotDict = annotation.GetPdfObject();
            PdfFreeTextAnnotation annotFT = annotation as PdfFreeTextAnnotation;
            PdfMarkupAnnotation annotMA = annotation as PdfMarkupAnnotation;

            //remove the AP
            annotDict.Remove(PdfName.AP);

            //set contents
            annotation.SetContents(contents);

            //*set default Style
            string contentDS = String.Format("font: Helvetica ,sans - serif {0}.00pt; color:{1}", fontSize, ColorTranslator.ToHtml(defaultColor));
            annotDict.Put(PdfName.DS, new PdfString(contentDS));

            //*set modification date
            annotation.SetDate(new PdfString(modificationDate));

            //*set Rich text
            annotMA.SetRichText(new PdfString(richText));

        }
        private static string _generateRichText(int fontSize, SysColor defaultColor, string content)
        {

            //get hex color
            string textColorHex = String.Format("#{0}{1}{2}", defaultColor.R.ToString("X2"), defaultColor.B.ToString("X2"), defaultColor.G.ToString("X2"));

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
        private static string _generateRichTextForRevBlock(int fontSize, SysColor defaultColor, string content)
        {

            //get hex color
            string textColorHex = String.Format("#{0}{1}{2}", defaultColor.R.ToString("X2"), defaultColor.B.ToString("X2"), defaultColor.G.ToString("X2"));

            //create richtext
            string richText = String.Format(
                "<?xml version=\"1.0\"?>" +
                "<body xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:xfa=\"http://www.xfa.org/schema/xfa-data/1.0/\" xfa:APIVersion=\"Acrobat:11.0.0\" xfa:spec=\"2.0.2\">" +
                "<p dir=\"ltr\">" +
                "<span style=\"line - height:4.23pt; text - align:left; font - size:6pt; font - style:normal; font - weight:normal; color:#000000;font-family:Helvetica\">&#x20; &#x0A;" +
                "</span>" +
                "</p>" +
                "<p dir=\"ltr\">" +
                "<span style=\"text-align:left;font-size:{0}pt;font-style:normal;font-weight:bold;color:{1};font-family:Helvetica\">{2}" +
                "</span>" +
                "</p>" +
                "</body>", fontSize, textColorHex, content);

            return richText;
        }
        private static void _modifyGeomAnnotColor(ref PdfAnnotation annot, SysColor newColor)
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
        private static bool addRevisionTitleblockOLD(string sourcePath, string destPath, string revision, string author)
        {
            PdfDocument pdfDoc = null;

            try
            {
                //get pdf for read and write
                pdfDoc = _getPdfForReadAndWrite(sourcePath, destPath);

                //get page
                PdfPage page = pdfDoc.GetFirstPage();

                //rectangles for each annot
                PdfArray[] tbRectangles = new PdfArray[] {
                new PdfArray(new float[] { 2805, 780, 2831, 794 }),
                new PdfArray(new float[] { 2836.2f,780,2889.8f,794 }),
                new PdfArray(new float[] { 2910,780,2935.5f,794 }),
                new PdfArray(new float[] { 2971.5f,780,2997.8f,794 }),
                new PdfArray(new float[] { 3028.6f,780,3055,794 }),
                new PdfArray(new float[] { 3070.6f,780,3279,794 }),
                new PdfArray(new float[] { 3304, 780, 3332, 794 })};
                

                PdfArray revBlocRect = new PdfArray(new float[] { 20802.53f, 686.377f, 3341.51f, 707.381f });

                //get date
                string time = DateTime.Now.ToString("dd/MM/yyyy", DateTimeFormatInfo.InvariantInfo);

                //contents for each annot
                string[] tbContents = new string[] { revision, time, author, "DP", "D4", "Updates as per FCR ", "RR" };

                for (int i = 0; i < tbContents.Length; i++)
                {
                    if (i == 1)
                    {
                        page.AddAnnotation(_createNewSimpleTypewriter(tbRectangles[i], tbContents[i], SysColor.Red, 10));
                    }
                    else
                    {
                        page.AddAnnotation(_createNewSimpleTypewriter(tbRectangles[i], tbContents[i], SysColor.Red, 11));
                    }
                }
                pdfDoc.Close();

                //confirm complete
                return true;
            }
            catch
            {
                //clean up
                if (pdfDoc != null) { pdfDoc.Close(); }

                //confirm incomplete
                return false;
            }

        }
        public static bool addRevisionTitleblock(string sourcePath, string destPath, string revision, string author, string checker)
        {
            PdfDocument pdfDoc = null;
            try
            {
                //get pdf for read and write
                pdfDoc = _getPdfForReadAndWrite(sourcePath, destPath);

                //get page
                PdfPage page = pdfDoc.GetFirstPage();

                //rectangles for titleblock annot
                float height = 22.65f;
                float freeY = _getNextFreeSpaceInRevBlock(page, height);
                PdfArray revBlocRect = new PdfArray(new float[] { 2802.53f, freeY, 3341.51f, freeY+height });

                //get date
                string time = DateTime.Now.ToString("dd/MM/yyyy", DateTimeFormatInfo.InvariantInfo);

                //contents for each annot
                string contents = string.Format("{0}     {1}          {2}                  {3}                D4                         Updated for FCR 000000                          RR", revision, time,author,checker);

                //add annot
                page.AddAnnotation(_createRevBlockTextbox(revBlocRect, contents, SysColor.Red, 10));

                //close documents
                pdfDoc.Close();

                //confirm complete
                return true;
            }
            catch
            {
                //clean up
                if (pdfDoc != null) { pdfDoc.Close(); }

                //confirm incomplete
                return false;
            }

        }
        private static PdfAnnotation _createNewTypewriter(PdfArray rectangle, string contents, SysColor defaultColor, string richText, string author, string creationDate, PdfNumber rotation, int fontSize)
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
            contentsBS.Put(PdfName.W, new PdfNumber(0));
            BS.Put(PdfName.BS, contentsBS);
            annotDict.Put(PdfName.BS, BS);

            //set the opacity
            annotDict.Put(PdfName.CA, new PdfNumber(1));

            //set content
            annotation.SetContents(contents);

            //set creation date 
            annotDict.Put(PdfName.CreationDate, new PdfString(creationDate));

            //*set default appearance
            annotDict.Put(PdfName.DA, new PdfString("   /FXF2 13 Tf"));
            AnnotationDefaultAppearance DA = new AnnotationDefaultAppearance();
            DA.SetColor(new DeviceRgb(defaultColor));
            DA.SetFont(StandardAnnotationFont.Helvetica);
            DA.SetFontSize(fontSize);
            annotFT.SetDefaultAppearance(DA);

            //*set default Style
            string contentDS = String.Format("font: Helvetica ,sans - serif {0}.00pt; color:{1}", fontSize, ColorTranslator.ToHtml(defaultColor));
            annotDict.Put(PdfName.DS, new PdfString(contentDS));

            //set flag
            annotation.SetFlag(4);

            //set intent
            annotMA.SetIntent(PdfName.FreeTextTypeWriter);

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
            annotDict.Put(PdfName.Subj, new PdfString("Typewriter"));

            //set subtype
            annotDict.Put(PdfName.Subtype, PdfName.FreeText);

            //set title
            annotation.SetTitle(new PdfString(author));

            //set type
            annotDict.Put(PdfName.Type, PdfName.Annot);

            return annotation;
        }
        private static PdfAnnotation _createNewSimpleTypewriter(PdfArray rectangle, string contents, SysColor defaultColor, int fontSize)
        {
            // make default values
            string author = Environment.UserName;
            string richText = _generateRichText(fontSize, defaultColor, contents);
            string creationDate = String.Format("D:{0}'00'", DateTime.Now.ToString("yyyyMMddHHmmsszz", DateTimeFormatInfo.InvariantInfo));
            PdfNumber rotation = new PdfNumber(0);

            //create annot
            return _createNewTypewriter(rectangle, contents, defaultColor, richText, author, creationDate, rotation, fontSize);
        }
        private static PdfAnnotation _createNewTextbox(PdfArray rectangle, string contents, SysColor defaultColor, string richText, string author, string creationDate, PdfNumber rotation, int fontSize)
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
        private static PdfAnnotation _createNewSimpleTextbox (PdfArray rectangle, string contents, SysColor defaultColor, int fontSize)
        {
            // make default values
            string author = Environment.UserName;
            string richText = _generateRichText(fontSize,defaultColor,contents);
            string creationDate = String.Format("D:{0}'00'", DateTime.Now.ToString("yyyyMMddHHmmsszz", DateTimeFormatInfo.InvariantInfo));
            PdfNumber rotation = new PdfNumber(0);

            //create annot
            return _createNewTextbox(rectangle, contents, defaultColor, richText, author, creationDate, rotation, fontSize);
        }
        private static PdfAnnotation _createRevBlockTextbox(PdfArray rectangle, string contents, SysColor defaultColor, int fontSize)
        {
            // make default values
            string author = Environment.UserName;
            string richText = _generateRichTextForRevBlock(fontSize, defaultColor, contents);
            string creationDate = String.Format("D:{0}'00'", DateTime.Now.ToString("yyyyMMddHHmmsszz", DateTimeFormatInfo.InvariantInfo));
            PdfNumber rotation = new PdfNumber(0);

            //create annot
            return _createNewTextbox(rectangle, contents, defaultColor, richText, author, creationDate, rotation, fontSize);
        }
        private static void _loopthroughimages(string sourcePath, string sigPath)
        {
            //get pdf for read and write
            PdfDocument pdfDoc = _getPdfForRead(sourcePath);

            //get page
            PdfPage page = pdfDoc.GetFirstPage();

            //ImageData imageData = ImageDataFactory.Create("logo.png");
            //pdfImage pdfImg = new pdfImage(imageData);

            for (int i = 1; i <= pdfDoc.GetNumberOfPdfObjects(); i++)
            {
                PdfObject obj = pdfDoc.GetPdfObject(i);

                if (obj != null && obj.IsStream())
                {
                    PdfDictionary pd = (PdfDictionary)obj;
                    if (pd.ContainsKey(PdfName.Subtype) && pd.Get(PdfName.Subtype).ToString() == "/Image")
                    {
                        //string test = "";
                    }
                }
            }
        }
        #endregion

        #region analise content
        public static string getNextRevisionFromAnnotations(string sourcePath)
        {
            //get pdf for read and write
            PdfDocument pdfDoc = _getPdfForRead(sourcePath);

            //get page
            PdfPage page = pdfDoc.GetFirstPage();

            //get annotation on page
            IList<PdfAnnotation> annotations = page.GetAnnotations();

            //get newest revision
            string newestRevision = "A.00";

            for (int i = 0; i < annotations.Count; i++)
            {
                if (annotations[i].GetSubtype() == PdfName.FreeText)
                {
                    //get content
                    PdfAnnotation annot = annotations[i] as PdfAnnotation;
                    string contents = annot.GetContents().ToString().TrimEnd(' ');

                    if (Regex.IsMatch(contents, RegexPattern.revision))
                    {
                        if (_isNewestRevision(contents, newestRevision))
                        {
                            newestRevision = contents;
                        }
                    }
                }
            }

            pdfDoc.Close();

            //+1 the current revision
            string output = _increaseRevision(newestRevision);

            return output;
        }
        public static string getNextRevisionFromDocumentName(string name)
        {
            string revNumber = name.Substring(name.LastIndexOf("-") + 1, 4);
            string newRevNumber = _increaseRevision(revNumber);
            return newRevNumber;
        }
        public static string getTextInRectangle(string sourcePath, float[] rectangle)
        {
            PdfDocument pdfDoc = null;
            try
            {
                //get pdf for read and write
                pdfDoc = _getPdfForRead(sourcePath);

                //get page
                PdfPage page = pdfDoc.GetFirstPage();

                //extract string
                PdfGeom.Rectangle rect = new PdfGeom.Rectangle(rectangle[0], rectangle[1], rectangle[2], rectangle[3]);
                TextRegionEventFilter regionFilter = new TextRegionEventFilter(rect);
                ITextExtractionStrategy strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), regionFilter);
                String str = PdfTextExtractor.GetTextFromPage(page, strategy);

                //close document
                pdfDoc.Close();

                //confirm complete
                return str;
            }
            catch
            {
                //clean up
                if (pdfDoc != null) { pdfDoc.Close(); }

                //confirm incomplete
                return "";
            }
        }
        private static string getTextInRectangle(PdfPage page, PdfGeom.Rectangle rectangle)
        {
            try
            {
                //extract string
                //PdfGeom.Rectangle rect = new PdfGeom.Rectangle(rectangle[0], rectangle[1], rectangle[2], rectangle[3]);
                TextRegionEventFilter regionFilter = new TextRegionEventFilter(rectangle);
                ITextExtractionStrategy strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), regionFilter);
                String str = PdfTextExtractor.GetTextFromPage(page, strategy);

                //confirm complete
                return str;
            }
            catch
            {
                //confirm incomplete
                return "";
            }
        }
        private static float _getNextFreeSpaceInRevBlock(PdfPage page, float height)
        {
            int rowCount = 8;
            float rowHeight = height;
            float minX = 2800f;
            float maxX = 3350f;
            float minY = 640.7f;
            float maxY = minY + (rowCount * rowHeight)+rowHeight;
            float freeY = minY;

            string extractedText = "";

            PdfGeom.Rectangle revBlock = new PdfGeom.Rectangle(minX, minY, maxX, maxY);

            try
            {
                //check for text
                for (int i = 0; i < rowCount; i++)
                {
                    //float[] rectnagle = new float[] { minX, minY + (rowHeight * i), maxX, minY + (rowHeight * (i + 1)) };
                    PdfGeom.Rectangle rectangle = new PdfGeom.Rectangle(minX, minY + (rowHeight * i), maxX - minX, rowHeight);
                    extractedText = getTextInRectangle(page, rectangle);
                    if (extractedText != "")
                    {
                        freeY = minY+(rowHeight*(i+1));
                        extractedText = "";
                    }
                    else
                    {
                        break;
                    }
                }

                //check for annots
                IList<PdfAnnotation> annotations = page.GetAnnotations();
                for (int i = 0; i < annotations.Count; i++)
                {
                    if (annotations[i].GetSubtype() == PdfName.FreeText)
                    {
                        //get content
                        PdfAnnotation annot = annotations[i] as PdfAnnotation;
                        PdfGeom.Rectangle annotRect = annot.GetRectangle().ToRectangle();
                        
                        //if its inside the revblock and higher than text
                        if (revBlock.Contains(annotRect) && annotRect.GetY()>freeY)
                        {
                            //get current row number based on freeY value
                            int rowNum = System.Convert.ToInt32((freeY - minY) / rowHeight);
                            //check in which row the annot is in
                            for (int j = rowNum; j < rowCount; j++)
                            {
                                PdfGeom.Rectangle rowrectanlge = new PdfGeom.Rectangle(minX, minY + (rowHeight * j), maxX - minX, rowHeight);
                                if (rowrectanlge.Contains(annotRect))
                                {
                                    freeY = minY + (rowHeight * (j + 1));
                                    break;
                                }
                            }
                        }
                    }
                }
                //confirm complete
                return freeY;
            }
            catch
            {
                //confirm incomplete
                return maxY-rowHeight;
            }
        }
        private static bool _isNewestRevision(string thisRev, string newestRev)
        {
            int[] thisRevArr = new int[2];
            thisRevArr[0] = char.ToUpper(thisRev[0]) - 64;
            thisRevArr[1] = int.Parse(thisRev.Substring(thisRev.LastIndexOf('.')+1, 2));

            int[] newestRevArr = new int[2];
            newestRevArr[0] = char.ToUpper(newestRev[0]) - 64;
            newestRevArr[1] = int.Parse(newestRev.Substring(newestRev.LastIndexOf('.')+1, 2));

            if (thisRevArr[0] > newestRevArr[0])
            {
                return true;
            }
            else if (thisRevArr[0].Equals(newestRevArr[0]))
            {
                if (thisRevArr[1] > newestRevArr[1])
                {
                    return true;
                }
            }
            return false;
        }
        private static string _increaseRevision (string newestRev)
        {
            int[] newestRevArr = new int[2];
            newestRevArr[0] = char.ToUpper(newestRev[0]) - 64;
            newestRevArr[1] = int.Parse(newestRev.Substring(newestRev.LastIndexOf('.')+1, 2));

            string letter = ((char)(newestRevArr[0] + 64)).ToString();
            string output = String.Format("{0}.{1,2:D2}", letter, newestRevArr[1]+1);
            
            return output;
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
        public static void testfromInternet(string destPath, string tempDestPath)
        {
            //get the drawing to copy from (soruce drawing)
            //PdfDocument sourceDoc = _getPdfForRead(sourcePath);

            //get the new drawing for read and temporary fro write (destination drawing)
            PdfDocument masterPdfDoc = _getPdfForReadAndWrite(destPath, tempDestPath);
            PdfPage masterDocPage = masterPdfDoc.GetFirstPage();

            //get annotations from source
            PdfPage sourcePage = masterPdfDoc.GetFirstPage();
            IList<PdfAnnotation> annotations = sourcePage.GetAnnotations();

            foreach (var annot in annotations)
            {
                if(annot.GetSubtype() == PdfName.FreeText)
                {
                    PdfDictionary annotAppDict = annot.GetAppearanceDictionary();
                    PdfAnnotationAppearance appearance = new PdfAnnotationAppearance(annotAppDict);

                    foreach (PdfName key in annotAppDict.KeySet())
                    {
                        PdfStream value = annotAppDict.GetAsStream(key);
                        PdfDictionary valueDict = annotAppDict.GetAsDictionary(key);

                        if (value != null)
                        {
                            var text = ExtractAnnotationText(value , masterPdfDoc);
                            PdfStream xObject = new PdfStream();
                        }
                    }
                }

            }

            //sourceDoc.Close();
            masterPdfDoc.Close();
        }
        private static String ExtractAnnotationText(PdfStream xObject, PdfDocument pfdDoc)
        {
            PdfResources resources = new PdfResources(xObject.GetAsDictionary(PdfName.Resources));
            resources.AddFont(pfdDoc, PdfFontFactory.CreateFont());
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
