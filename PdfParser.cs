using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows;
using System.Numerics;
using System.Text.RegularExpressions;
using SysColor = System.Drawing.Color;
using iText.Kernel.Colors;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Xobject;
using iText.Kernel.Font;
using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Pdf.Annot;
using iText.Kernel.Pdf.Annot.DA;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Filter;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Layout;
using pdfElement = iText.Layout.Element;
using PdfGeom = iText.Kernel.Geom;

namespace PdfParser
{



    public class PdfParser
    {
        private static class RegexPattern
        {
            public static String annotRev { get { return @"^[A-Z]\.[0-9][0-9]\s*$"; } }
            public static String annotFcr { get { return @"^[Ff][Cc][Rr]\s*[0-9][0-9][0-9][0-9][0-9][0-9]\s*$"; } }
            public static String nnbName { get { return @"^HPC-UK1226-U\d-\w\w\w-\w\w\w-\d\d\d\d\d\d$"; } }
            public static String nnbFullRev { get { return @"[A-Z]\.[0-9][0-9]"; } }
            public static String nnbLetterRev { get { return @"[A-Z]"; } }
            public static String nnbNameWithFullRev { get { return @"HPC-UK1226-U\d-\w\w\w-\w\w\w-\d\d\d\d\d\d-\w\.\d\d"; } }
            public static String nnbNameWithLetterRev { get { return @"HPC-UK1226-U\d-\w\w\w-\w\w\w-\d\d\d\d\d\d-\w"; } }
        }

        private class RevCellDims
        {
            public int rowCount { get; }
            //rev box
            public float revBoxX { get; }
            public float revBoxWidth { get; }
            public float revBoxY { get; }
            public float rowHeight { get; }
            public float rowSpacing { get; }

            public float cellMargin { get; }
            public float authCellX { get; }
            public float authCellWidth { get; }
            public float chckCellX { get; }
            public float chckCellWidth { get; }
            public float apprCellX { get; }
            public float apprCellWidth { get; }

            public RevCellDims( int rowCount, float revBoxX, float revBoxWidth, float revBoxY, float rowHeight, float rowSpacing,
                                float cellMargin, float authCellX, float authCellWidth, float chckCellX, float chckCellWidth,
                                float apprCellX, float apprCellWidth)
            {
                this.rowCount = rowCount;
                this.revBoxX = revBoxX;
                this.revBoxWidth = revBoxWidth;
                this.revBoxY = revBoxY;
                this.rowHeight = rowHeight;
                this.rowSpacing = rowSpacing;
                this.cellMargin = cellMargin;
                this.authCellX = authCellX;
                this.authCellWidth = authCellWidth;
                this.chckCellX = chckCellX;
                this.chckCellWidth = chckCellWidth;
                this.apprCellX = apprCellX;
                this.apprCellWidth = apprCellWidth;
            }
        }

        private static class RevCellStandardOLD
        {
            public static int rowCount { get { return 8; } }
            //rev box
            public static float revBoxX { get { return 2803.63f; } }
            public static float revBoxWidth { get { return 538.22f; } }
            public static float revBoxY { get { return 640.76f; } }
            public static float rowHeight { get { return 22.5f; } }
            public static float rowSpacing { get { return 0.271f; } }

            //sigs
            public static float cellMargin { get { return 2.5f; } }
            public static float authCellX { get { return 2891.6f; } }
            public static float authCellWidth { get { return 61.9f; } }
            public static float chckCellX { get { return 2953.87f; } }
            public static float chckCellWidth { get { return 61.9f; } }
            public static float apprCellX { get { return 3280f; } }
            public static float apprCellWidth { get { return 61.74f; } }
        }
        private static class RevCellSmallOLD
        {
            public static int rowCount { get { return 8; } }
            //rev box
            public static float revBoxX { get { return 2790.18f; } }
            public static float revBoxWidth { get { return 531.79f; } }
            public static float revBoxY { get { return 647.393f; } }
            public static float rowHeight { get { return 22.12f; } }
            public static float rowSpacing { get { return 0.271f; } }

            //sigs
            public static float cellMargin { get { return 2.5f; } }
            public static float authCellX { get { return 2876.92f; } }
            public static float authCellWidth { get { return 61.2f; } }
            public static float chckCellX { get { return 2938.56f; } }
            public static float chckCellWidth { get { return 61.13f; } }
            public static float apprCellX { get { return 3260.56f; } }
            public static float apprCellWidth { get { return 61.2f; } }

        }

        //Data for locations and sizes of revision block and new annotations
        //normal when margins are set correctly by drawing originator
        RevCellDims revDimsNormal = new RevCellDims(
        rowCount: 8,
        revBoxX: 2803.63f,
        revBoxWidth: 538.22f,
        revBoxY: 640.76f,
        rowHeight: 22.5f,
        rowSpacing: 0.271f,
        cellMargin: 2.5f,
        authCellX: 2891.6f,
        authCellWidth: 61.9f,
        chckCellX: 2953.87f,
        chckCellWidth: 61.9f,
        apprCellX: 3280f,
        apprCellWidth: 61.74f
        );

        //smaller when margins are set bigger than normal by drawing originator
        RevCellDims revDimsSmall = new RevCellDims(
        rowCount: 8,
        revBoxX: 2790.07f,
        revBoxWidth: 531.79f,
        revBoxY: 647.393f,
        rowHeight: 22.25f,
        rowSpacing: 0.25f,
        cellMargin: 2.5f,
        authCellX: 2876.92f,
        authCellWidth: 61.2f,
        chckCellX: 2938.56f,
        chckCellWidth: 61.13f,
        apprCellX: 3260.56f,
        apprCellWidth: 61.2f
        );

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

        #region ###### copy content
        public static bool copyAnnotsByColorRed(string sourcePath, string destPath, string tempDestPath)
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
                //_filterAnnotationsByColor(ref annotations, filterColor);

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
        private static void _filterAnnotationsByRegex(ref IList<PdfAnnotation> annots, string regexPattern)
        {
            for (int i = annots.Count - 1; i >= 0; i--)
            {
                PdfAnnotation annot = annots[i] as PdfAnnotation;
                string annotText = annot.GetContents().ToString();

                if (!Regex.IsMatch(annotText, regexPattern))
                {
                    annots.RemoveAt(i);
                }
            }
        }
        private static void _filterAnnotationsByType(ref IList<PdfAnnotation> annots, PdfName annotTypeName)
        {
            for (int i = annots.Count - 1; i >= 0; i--)
            {
                PdfAnnotation annot = annots[i] as PdfAnnotation;

                if (annot.GetSubtype() != annotTypeName)
                {
                    annots.RemoveAt(i);
                }
            }
        }
        private static void _filterAnnotationsByType(ref IList<PdfAnnotation> annots, List<PdfName> annotTypeNames)
        {
            for (int i = annots.Count - 1; i >= 0; i--)
            {
                PdfAnnotation annot = annots[i] as PdfAnnotation;

                foreach(PdfName typeName in annotTypeNames)
                {
                    bool found = false;
                    if (annot.GetSubtype() == typeName)
                    {
                        found = true;
                    }
                    if (!found)
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

        #region ####### change content
        public static bool changeTextboxContentsByRegex (string sourcePath, string tempPath, string newContent)
        {
            PdfDocument pdfDoc = null;
            try
            {
                //get pdf for read and write
                pdfDoc = _getPdfForReadAndWrite(sourcePath, tempPath);

                //get page
                PdfPage page = pdfDoc.GetFirstPage();

                //get annotation on page
                IList<PdfAnnotation> annots = page.GetAnnotations();

                //filter the annotations to only include red annots (most recent annots)
                _filterAnnotationsByColor(ref annots, SysColor.Red);

                //change textboxes using the correct function depending on type of content passed in
                if (Regex.IsMatch(newContent, RegexPattern.annotRev)) { _changeTextboxContentUsingRegex(ref annots, newContent, RegexPattern.annotRev ); }
                else if(Regex.IsMatch(newContent,RegexPattern.annotFcr)) { _changeTextboxContentUsingRegex(ref annots, newContent, RegexPattern.annotFcr); }

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
                _changeAnnotsUsingColor(ref annotations, existingColor, newColor);

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
        private static void _changeAnnotsUsingColor(ref IList<PdfAnnotation> annots, SysColor existingColor, SysColor newColor)
        {

            //filter annotations by color
            _filterAnnotationsByColor(ref annots, existingColor);

            for (int i = 0; i < annots.Count; i++)
            {
                PdfAnnotation annot = annots[i] as PdfAnnotation;
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
        private static void _changeTextboxContentUsingRegex(ref IList<PdfAnnotation> annots, string newContent, string regexPattern)
        {
            for (int i = 0; i < annots.Count; i++)
            {
                if (annots[i].GetSubtype() == PdfName.FreeText)
                {
                    //get content
                    PdfAnnotation annot = annots[i] as PdfAnnotation;
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
        private static void _modifyTextBox(ref PdfAnnotation annot, string contents, SysColor defaultColor, int fontSize, string richText, string modificationDate)
        {
            PdfDictionary annotDict = annot.GetPdfObject();
            PdfFreeTextAnnotation annotFT = annot as PdfFreeTextAnnotation;
            PdfMarkupAnnotation annotMA = annot as PdfMarkupAnnotation;

            //set the AP
            PdfAnnotationAppearance annotAP = generateFreetextAppearance(annot);
            annot.SetNormalAppearance(annotAP);

            //get length
            PdfStream value = annot.GetAppearanceDictionary().GetAsStream(PdfName.N);

            //set contents
            annot.SetContents(contents);

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
            annot.SetDate(new PdfString(modificationDate));

            //*set Rich text
            annotMA.SetRichText(new PdfString(richText));

        }
        private static void _modifyTypewriter(ref PdfAnnotation annot, string contents, SysColor defaultColor, int fontSize, string richText, string modificationDate)
        {
            PdfDictionary annotDict = annot.GetPdfObject();
            PdfFreeTextAnnotation annotFT = annot as PdfFreeTextAnnotation;
            PdfMarkupAnnotation annotMA = annot as PdfMarkupAnnotation;

            //remove the AP
            PdfAnnotationAppearance annotAP = generateFreetextAppearance(annot);
            annot.SetNormalAppearance(annotAP);

            //get length
            PdfStream value = annot.GetAppearanceDictionary().GetAsStream(PdfName.N);

            //set contents
            annot.SetContents(contents);

            //*set default Style
            string contentDS = String.Format("font: Helvetica ,sans - serif {0}.00pt; color:{1}", fontSize, ColorTranslator.ToHtml(defaultColor));
            annotDict.Put(PdfName.DS, new PdfString(contentDS));

            //*set modification date
            annot.SetDate(new PdfString(modificationDate));

            //*set Rich text
            annotMA.SetRichText(new PdfString(richText));

        }
        private static void _modifyGeomAnnotColor(ref PdfAnnotation annot, SysColor newColor)
        {
            PdfDictionary annotDict = annot.GetPdfObject();

            //remove the AP
            PdfAnnotationAppearance annotAP = generateGeomAppearance(annot);
            annot.SetNormalAppearance(annotAP);

            //set color
            float[] floatRGB = new float[3];
            for (int j = 0; j < 2; j++)
            {
                floatRGB[j] = (float)System.Convert.ToDouble(newColor.R) / 255;
            }
            annot.SetColor(floatRGB);
        }
        private static PdfAnnotationAppearance generateGeomAppearance(PdfAnnotation annot)
        {
            PdfGeom.Rectangle annotRectangle = annot.GetRectangle().ToRectangle();
            PdfDocument pdfDoc = annot.GetPage().GetDocument();

            PdfFormXObject appearObj = new PdfFormXObject(annotRectangle);
            // add matrix
            appearObj.Put(PdfName.Matrix, new PdfArray(new float[] { 1, 0, 0, 1, -annotRectangle.GetX(), -annotRectangle.GetY() }));
            //add filter
            appearObj.Put(PdfName.Filter, PdfName.FlateDecode);
            //add form type
            appearObj.Put(PdfName.FormType, new PdfNumber(1));
            // flush object
            appearObj.MakeIndirect(pdfDoc);
            //add length
            //PdfOutputStream outStream = appearObj.GetPdfObject().GetOutputStream();
            //appearObj.Put(PdfName.Length, new PdfNumber(outStream.Length));

            return new PdfAnnotationAppearance(appearObj.GetPdfObject());
        }
        private static PdfAnnotationAppearance generateFreetextAppearance(PdfAnnotation annot)
        {
            PdfGeom.Rectangle annotRectangle = annot.GetRectangle().ToRectangle();
            PdfDocument pdfDoc = annot.GetPage().GetDocument();

            PdfFormXObject appearObj = new PdfFormXObject(annotRectangle);
            // add matrix
            appearObj.Put(PdfName.Matrix, new PdfArray(new float[] { 1, 0, 0, 1, -annotRectangle.GetX(), -annotRectangle.GetY() }));
            //add filter
            appearObj.Put(PdfName.Filter, PdfName.FlateDecode);
            //add form type
            appearObj.Put(PdfName.FormType, new PdfNumber(1));
            //add resources
            PdfResources resources = new PdfResources();
            resources.AddFont(pdfDoc, PdfFontFactory.CreateFont(iText.IO.Font.Constants.StandardFonts.HELVETICA));
            appearObj.Put(PdfName.Resources, resources.GetPdfObject());

            // flush object
            appearObj.MakeIndirect(pdfDoc);

            //add length
            //PdfOutputStream outStream = appearObj.GetPdfObject().GetOutputStream();
            //appearObj.Put(PdfName.Length, new PdfNumber(outStream.Length));
            

            return new PdfAnnotationAppearance(appearObj.GetPdfObject());
        }
        #endregion

        #region add content
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
        public static bool addRevisionTitleblockOLD2(string sourcePath, string destPath, string revision, string author, string checker, string reason, string authSigPath = "", string checkerSigPath = "", string approverSigPath = "")
        {

            //Data for locations and sizes of revision block and new annotations
            //normal when margins are set correctly by drawing originator
            RevCellDims RevCellNormal = new RevCellDims(
            rowCount: 8,
            revBoxX: 2803.63f,
            revBoxWidth: 538.22f,
            revBoxY: 640.76f,
            rowHeight: 22.5f,
            rowSpacing: 0.271f,
            cellMargin: 2.5f,
            authCellX: 2891.6f,
            authCellWidth: 61.9f,
            chckCellX: 2953.87f,
            chckCellWidth: 61.9f,
            apprCellX: 3280f,
            apprCellWidth: 61.74f
            );

            //smaller when margins are set bigger than normal by drawing originator
            RevCellDims RevCellSmaller = new RevCellDims(
            rowCount: 8,
            revBoxX: 2790.07f,
            revBoxWidth: 531.79f,
            revBoxY: 647.393f,
            rowHeight: 22.25f,
            rowSpacing: 0.25f,
            cellMargin: 2.5f,
            authCellX: 2876.92f,
            authCellWidth: 61.2f,
            chckCellX: 2938.56f,
            chckCellWidth: 61.13f,
            apprCellX: 3260.56f,
            apprCellWidth: 61.2f
            );


            PdfDocument pdfDoc = null;
            try
            {
                //get pdf for read and write
                pdfDoc = _getPdfForReadAndWrite(sourcePath, destPath);

                //get page
                PdfPage page = pdfDoc.GetFirstPage();

                //check if drawing has normal margins
                bool pageMarginsAreNormal = _checkIfPageMarginsAreNormal(page);

                //select correct revision block positions and dimensions
                RevCellDims revBlockDims = null;
                if (pageMarginsAreNormal) { revBlockDims = RevCellNormal; }
                else { revBlockDims = RevCellSmaller; }

                //get free space in revtitleblock depeding on margins
                float freeY = _getNextFreeSpaceInRevBlock(page, revBlockDims);

                //add rev titleblock
                //get rectangle
                PdfArray revBlocRect = new PdfArray(new float[] { revBlockDims.revBoxX, freeY, revBlockDims.revBoxX + revBlockDims.revBoxWidth, freeY + revBlockDims.rowHeight });
                //get date
                string time = DateTime.Now.ToString("dd/MM/yy", DateTimeFormatInfo.InvariantInfo);
                //contents for each annot
                string contents = string.Format("{0}   {1}         {2}               {3}             D4              {4}                    RR", revision, time, author, checker, reason);
                //add annot
                page.AddAnnotation(_createRevBlockTextbox(revBlocRect, contents, SysColor.Red, 12));

                //add sigs
                if (authSigPath != "")
                {
                    float cellX = revBlockDims.authCellX;
                    float width = revBlockDims.authCellWidth;
                    float margin = revBlockDims.cellMargin;
                    float height = revBlockDims.rowHeight;

                    PdfGeom.Rectangle authSigRect = new PdfGeom.Rectangle(cellX + margin, freeY + height + margin, width - (2 * margin), freeY + (height * 2) - (2 * margin));
                    addSignaturesToRevision(pdfDoc, authSigRect, authSigPath);
                }
                if (checkerSigPath != "")
                {
                    float cellX = revBlockDims.chckCellX;
                    float width = revBlockDims.chckCellWidth;
                    float margin = revBlockDims.cellMargin;
                    float height = revBlockDims.rowHeight;

                    PdfGeom.Rectangle checkerSigRect = new PdfGeom.Rectangle(cellX + margin, freeY + height + margin, width - (2 * margin), freeY + (height * 2) - (2 * margin));
                    addSignaturesToRevision(pdfDoc, checkerSigRect, checkerSigPath);
                }
                if (approverSigPath != "")
                {
                    float cellX = revBlockDims.apprCellX;
                    float width = revBlockDims.apprCellWidth;
                    float margin = revBlockDims.cellMargin;
                    float height = revBlockDims.rowHeight;

                    PdfGeom.Rectangle approverSigRect = new PdfGeom.Rectangle(cellX + margin, freeY + height + margin, width - (2 * margin), freeY + (height * 2) - (2 * margin));
                    addSignaturesToRevision(pdfDoc, approverSigRect, approverSigPath);
                }

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
        public bool addRevisionTitleblock(string sourcePath, string destPath, string revision, string author, string checker, string reason, string authSigPath = "", string checkerSigPath = "", string approverSigPath = "")
        {
            PdfDocument pdfDoc = null;
            try
            {
                //get pdf for read and write
                pdfDoc = _getPdfForReadAndWrite(sourcePath, destPath);

                //get page
                PdfPage page = pdfDoc.GetFirstPage();

                //select correct revision block positions and dimensions
                RevCellDims revBlockDims = _getRevisionBlockDimensions(page);

                //get free space in revtitleblock depeding on margins
                float freeY = _getNextFreeSpaceInRevBlock(page, revBlockDims);

                //add rev titleblock
                //get rectangle
                PdfArray revBlocRect = new PdfArray(new float[] { revBlockDims.revBoxX, freeY, revBlockDims.revBoxX + revBlockDims.revBoxWidth, freeY+ revBlockDims.rowHeight });
                //get date
                string time = DateTime.Now.ToString("dd/MM/yy", DateTimeFormatInfo.InvariantInfo);
                //contents for each annot
                string contents = string.Format("{0}   {1}         {2}               {3}             D4              {4}                    RR", revision, time, author, checker, reason);
                //add annot
                page.AddAnnotation(_createRevBlockTextbox(revBlocRect, contents, SysColor.Red, 12));

                //add sigs
                if (authSigPath != "") 
                {
                    float cellX = revBlockDims.authCellX;
                    float width = revBlockDims.authCellWidth;
                    float margin = revBlockDims.cellMargin;
                    float height = revBlockDims.rowHeight;

                    PdfGeom.Rectangle authSigRect = new PdfGeom.Rectangle(cellX + margin, freeY + height+margin, width-(2*margin), freeY + (height * 2) - (2*margin));
                    addSignaturesToRevision(pdfDoc, authSigRect, authSigPath);
                }
                if (checkerSigPath != "")
                {
                    float cellX = revBlockDims.chckCellX;
                    float width = revBlockDims.chckCellWidth;
                    float margin = revBlockDims.cellMargin;
                    float height = revBlockDims.rowHeight;

                    PdfGeom.Rectangle checkerSigRect = new PdfGeom.Rectangle(cellX + margin, freeY + height + margin, width - (2 * margin), freeY + (height * 2) - (2 * margin));
                    addSignaturesToRevision(pdfDoc, checkerSigRect, checkerSigPath);
                }
                if (approverSigPath != "")
                {
                    float cellX = revBlockDims.apprCellX;
                    float width = revBlockDims.apprCellWidth;
                    float margin = revBlockDims.cellMargin;
                    float height = revBlockDims.rowHeight;

                    PdfGeom.Rectangle approverSigRect = new PdfGeom.Rectangle(cellX + margin, freeY + height + margin, width - (2 * margin), freeY + (height * 2) - (2 * margin));
                    addSignaturesToRevision(pdfDoc, approverSigRect, approverSigPath);
                }

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
        private static bool addSignaturesToRevision(PdfDocument pdfDoc, PdfGeom.Rectangle sigPosition, string sigPath)
        {
            Document doc = new Document(pdfDoc);
            
            ImageData imgData = ImageDataFactory.Create(sigPath);
            pdfElement.Image img = new pdfElement.Image(imgData, sigPosition.GetX(), sigPosition.GetY(), sigPosition.GetWidth());
            doc.Add(img);
            return true;
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
                "<span style=\"line - height:4.23pt; text - align:left; font - size:4pt; font - style:normal; font - weight:normal; color:{3};font-family:Helvetica\">&#x20; &#x0A;" +
                "</span>" +
                "</p>" +
                "<p dir=\"ltr\">" +
                "<span style=\"text-align:left;font-size:{0}pt;font-style:normal;font-weight:bold;color:{1};font-family:Helvetica\">{2}" +
                "</span>" +
                "</p>" +
                "</body>", fontSize, textColorHex, content, textColorHex);

            return richText;
        }
        #endregion

        #region analyse content
        public string getCurrentRevisionFromDocument(string sourcePath)
        {
            //get pdf for read and write
            PdfDocument pdfDoc = _getPdfForRead(sourcePath);

            //get page
            PdfPage page = pdfDoc.GetFirstPage();

            //get newest revision
            string newestRevision = "A";

            //check if margins are normal
            RevCellDims revBlockDims = _getRevisionBlockDimensions(page);

            //check text in the revision block
            newestRevision = _getCurrentRevisionFromRevisionBlock(page, revBlockDims, newestRevision);

            //get annotation on page
            IList<PdfAnnotation> annots = page.GetAnnotations();

            //check rev annotations in drawing
            newestRevision = _getCurrentRevisionFromAnnots(annots, newestRevision);

            //close docs
            pdfDoc.Close();

            return newestRevision;
        }
        public static string getNextRevisionFromNameOLD(string name)
        {
            //get last portion of the name
            string lastNamePortion = name.Substring(name.LastIndexOf("-") + 1);

            bool hasNumericalRev = Regex.IsMatch(name, RegexPattern.nnbFullRev);
            bool hasLetterRev = Regex.IsMatch(lastNamePortion, RegexPattern.nnbLetterRev);

            if (hasNumericalRev)
            {
                string rev = Regex.Match(name, RegexPattern.nnbFullRev).Value;
                string newRevNumber = _increaseNumericalRevision(rev);
                return newRevNumber;
            }
            else if(hasLetterRev)
            {
                string newRevNumber = lastNamePortion + ".01";
                return newRevNumber;
            }

            //else return nothing
            return "";
        }
        public static string getNextNumericalRev(string currentRevision)
        {
            return _increaseNumericalRevision(currentRevision);
        }
        public static string getNextLetterRev(string currentRevision)
        {
            return _increaseLetterRevision(currentRevision);
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
        public static void loopthroughimages(string sourcePath, string sigPath)
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
        private static bool _checkIfPageMarginsAreNormal(PdfPage page)
        {
            float X = 5.5f;
            float Y = 15;
            float W = 10f;
            float H = 75f;
            string extractedText = "";

            PdfGeom.Rectangle rectangle = new PdfGeom.Rectangle(X, Y, W, H);
            extractedText = _getTextInRectangle(page, rectangle);
            if (extractedText == "")
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        private static string _getCurrentRevisionFromAnnots(IList<PdfAnnotation> annots, string baseRevision)
        {
            string newestRevision = baseRevision;

            for (int i = 0; i < annots.Count; i++)
            {
                if (annots[i].GetSubtype() == PdfName.FreeText)
                {
                    //get content
                    PdfAnnotation annot = annots[i] as PdfAnnotation;
                    string contents = annot.GetContents().ToString().TrimEnd(' ');

                    if (Regex.IsMatch(contents, RegexPattern.annotRev))
                    {
                        if (_isNewestRevision(contents, newestRevision))
                        {
                            newestRevision = contents;
                        }
                    }
                }
            }

            return newestRevision;
        }
        private static string _getCurrentRevisionFromRevisionBlock(PdfPage page, RevCellDims dims, string baseRevision)
        {
            int rowCount = dims.rowCount;
            float rowHeight = dims.rowHeight;
            float minX = dims.revBoxX;
            float maxX = dims.revBoxX + dims.revBoxWidth;
            float minY = dims.revBoxY;
            float maxY = minY + (rowCount * rowHeight) + rowHeight;
            float freeY = minY;

            string extractedText = "";

            //get newest revision
            string newestRevision = baseRevision;

            try
            {
                //check for text
                for (int i = 0; i < rowCount; i++)
                {
                    PdfGeom.Rectangle rectangle = new PdfGeom.Rectangle(minX, minY + (rowHeight * i), maxX - minX, rowHeight);
                    extractedText = _getTextInRectangle(page, rectangle);
                    //if extracted text matches pattern, extract the revision 
                    //check if extracted revision is latest and save
                }

                //complete
                return newestRevision;
            }
            catch
            {
                //incomplete
                return newestRevision;
            }
        }
        private static string _getTextInRectangle(PdfPage page, PdfGeom.Rectangle rectangle)
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
        private static float _getNextFreeSpaceInRevBlock(PdfPage page, RevCellDims dimensions)
        {
            int rowCount = dimensions.rowCount;
            float rowHeight = dimensions.rowHeight;
            float minX = dimensions.revBoxX;
            float maxX = dimensions.revBoxX + dimensions.revBoxWidth;
            float minY = dimensions.revBoxY;
            float maxY = minY + (rowCount * rowHeight)+rowHeight;
            float freeY = minY;

            string extractedText = "";

            // extents of revision block on drawing. Bottom of rectangle = bottom corner of revision A row + 5 point margin top, bottm and sides
            // used to filter out annots that are positioned within the revision block area.
            PdfGeom.Rectangle revBlock = new PdfGeom.Rectangle(minX - 5, minY - 5, maxX - minX + 10, maxY + 10);

            try
            {
                //check for text
                for (int i = 0; i < rowCount; i++)
                {
                    //float[] rectnagle = new float[] { minX, minY + (rowHeight * i), maxX, minY + (rowHeight * (i + 1)) };
                    PdfGeom.Rectangle rectangle = new PdfGeom.Rectangle(minX, minY + (rowHeight * i), maxX - minX, rowHeight);
                    extractedText = _getTextInRectangle(page, rectangle);
                    if (extractedText != "")
                    {
                        freeY = minY + (rowHeight * (i + 1)) + (dimensions.rowSpacing * i);
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
                        if (revBlock.Contains(annotRect) && (annotRect.GetY()+annotRect.GetHeight())>freeY)
                        {
                            //get current row number based on freeY value
                            int rowNum = System.Convert.ToInt32((freeY - minY) / rowHeight);
                            //check in which row the annot is in
                            for (int j = rowNum; j < rowCount; j++)
                            {
                                // extents of reach row that could contain the annotation. Bottom of rectangle = bottom corner of revision given row + 5 point margin top, bottm and sides
                                PdfGeom.Rectangle rowrectanlge = new PdfGeom.Rectangle(minX-5, minY + (rowHeight * j)-5, maxX - minX+10, rowHeight+10);
                                if (rowrectanlge.Contains(annotRect))
                                {
                                    freeY = minY + (rowHeight * (j + 1)) + (dimensions.rowSpacing*j);
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
        private RevCellDims _getRevisionBlockDimensions(PdfPage page)
        {
            RevCellDims revBlockDims;

            float X = 5.5f;
            float Y = 15;
            float W = 10f;
            float H = 75f;
            string extractedText = "";

            PdfGeom.Rectangle rectangle = new PdfGeom.Rectangle(X, Y, W, H);
            extractedText = _getTextInRectangle(page, rectangle);
            if (extractedText == "")
            {
                return revDimsSmall;
            }
            else
            {
                return revDimsNormal;
            }
        }
        private static bool _isNewestRevision(string revToCheck, string baseRev)
        {

            //accepts single letter ("A") and numerical ("A.01") formats

            //THIS REVISION
            int[] revToCheckArr = _translateRevisionTextToInt(revToCheck);

            // NEWEST REVISION
            int[] baseRevArr = _translateRevisionTextToInt(baseRev);


            //compare LETTERS
            //if letter is higher, then must be newer.
            if (revToCheckArr[0] > baseRevArr[0])
            {
                return true;
            }

            //comapre NUMBERS
            //if newer has numerical but base doesnt
            if(revToCheckArr.Length > baseRevArr.Length)
            {
                return true;
            }

            //(both have numbers) and letter is the same
            if((baseRevArr.Length == revToCheckArr.Length && baseRev.Length > 1) && revToCheckArr[0].Equals(baseRevArr[0]))
            {
                //if numbers are higher it must be newer
                if (revToCheckArr[1] > baseRevArr[1])
                {
                    return true;
                }
            }

            //otherwise, ignore as its either lower or the same
            return false;
        }
        private static string _increaseNumericalRevision (string thisRev)
        {
            //if its just a letter then add 00 numericals
            if(thisRev.Length == 1)
            {
                thisRev = thisRev + ".00";
            }

            //THIS REVISION
            int[] thisRevArr = _translateRevisionTextToInt(thisRev);

            //increase numerical revision by one
            thisRevArr[1] = thisRevArr[1] + 1;
            
            return _translateRevIntToText(thisRevArr);
        }
        private static string _increaseLetterRevision(string thisRev)
        {
            //THIS REVISION
            int[] thisRevArr = _translateRevisionTextToInt(thisRev);

            //increase letter revision by one
            thisRevArr[0] = thisRevArr[0] + 1;

            return _translateRevIntToText(thisRevArr);
        }
        private static int[] _translateRevisionTextToInt(string revString)
        {
            //accepts single letter ("A") and numerical ("A.01") formats

            //0 position denotes revision letter. Parse char into number for numerical comparison
            //optional - 1 position denotes numerical revision using last two characters. (assuming numerical revision is always two digits 01-09, 10, 11..)

            // is length includes numerical then include in translation
            if (revString.Length >= 4)
            {
                int[] revArr = new int[2];
                revArr[0] = char.ToUpper(revString[0]) - 64;
                revArr[1] = int.Parse(revString.Substring(revString.LastIndexOf('.') + 1, 2));
                return revArr;
            }
            else
            {
                int[] revArr = new int[1];
                revArr[0] = char.ToUpper(revString[0]) - 64;
                return revArr;
            }
        }
        private static string _translateRevIntToText(int[] thisRevArr)
        {
            if(thisRevArr.Length == 1)
            {
                //translate letter
                string letter = ((char)(thisRevArr[0] + 64)).ToString();

                return letter;
            }
            else
            {
                //translate letter
                string letter = ((char)(thisRevArr[0] + 64)).ToString();
                //translate number
                int number = thisRevArr[1];

                //create string 
                return String.Format("{0}.{1:D2}", letter, number);
            }
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
                if (annot.GetSubtype() == PdfName.FreeText)
                {
                    PdfDictionary annotAppDict = annot.GetAppearanceDictionary();
                    PdfAnnotationAppearance appearance = new PdfAnnotationAppearance(annotAppDict);

                    foreach (PdfName key in annotAppDict.KeySet())
                    {
                        PdfStream value = annotAppDict.GetAsStream(key);
                        PdfDictionary valueDict = annotAppDict.GetAsDictionary(key);
                        PdfXObject appearObj = new PdfFormXObject(annot.GetRectangle().ToRectangle());

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

        private static void test()
        {
            double[] values = new double[2] { 0.5, 0.5 };
            float[] fValues = new float[2] { 0.5f, 0.5f };
            float[] fValues2 = new float[2] { 1f, 1f };

            Vector<float> newV = new Vector<float>(fValues);
            Vector<float> newV2 = new Vector<float>(fValues2);

            Vector<float> addedV = Vector.Subtract<float>(newV, newV2);
        }

        public static void regexExtractionStratTest(string path)
        {
            PdfReader reader = new PdfReader(path);
            PdfDocument pdfDoc = new PdfDocument(reader);
            Document doc = new Document(pdfDoc);
            PdfPage page = pdfDoc.GetPage(1);


            string pattern = "UK PROTECT";

            RegexBasedLocationExtractionStrategy strategy = new RegexBasedLocationExtractionStrategy(pattern);
            PdfCanvasProcessor parser = new PdfCanvasProcessor(strategy);
            parser.ProcessPageContent(page);
            //parser.ProcessPageContent(page);
            ICollection<IPdfTextLocation> locations = strategy.GetResultantLocations();
            IList<IPdfTextLocation> list = (IList<IPdfTextLocation>)locations;
        }

        #endregion
    }
}
