using System;
using System.Collections.Generic;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser.Filter;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Annot;
using iText.Layout;
using iText.Layout.Element;
using System.Drawing;
using System.Text.RegularExpressions;




namespace PdfParser
{
    public static class PdfParser
    {
        //###### General methods

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

        //###### copy annotations by color

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
            int[] redRGB = new int[] { 255, 0, 0 };
            filterAnnotationsByColor(ref annotations,redRGB);

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
        public static void filterAnnotationsByColor(ref IList<PdfAnnotation> annots, int[] filterColor)
        {
            for (int i = annots.Count - 1; i >= 0; i--)
            {
                PdfAnnotation annot = annots[i] as PdfAnnotation;
                int[] annotColor = getAnnotationColors(annot);
                if (annotColor!= null)
                {
                    if (!checkIfColorIsSimilar(filterColor, annotColor, 20))
                    {
                        annots.RemoveAt(i);
                    }
                }

            }
        }
        public static void excludeAnnotationsByColor(ref IList<PdfAnnotation> annots, int[] filterColor)
        {
            for (int i = annots.Count - 1; i >= 0; i--)
            {
                PdfAnnotation annot = annots[i] as PdfAnnotation;
                int[] annotColor = getAnnotationColors(annot);
                if (annotColor != null)
                {
                    if (checkIfColorIsSimilar(filterColor, annotColor, 20))
                    {
                        annots.RemoveAt(i);
                    }
                }

            }
        }
        private static int[] getAnnotationColors(PdfAnnotation annot)
        {
            int[] colors = null;
            if (annot.GetSubtype() == PdfName.FreeText)
            {
                PdfFreeTextAnnotation freeText = annot as PdfFreeTextAnnotation;

                PdfString defaultApperance = freeText.GetDefaultAppearance();
                string[] daText = defaultApperance.ToString().Split(' ');

                PdfString defaultStyle = freeText.GetDefaultStyleString();
                string[] dsText = defaultStyle.ToString().Split(';');

                PdfMarkupAnnotation markupAnnot = annot as PdfMarkupAnnotation;
                PdfObject richText = markupAnnot.GetRichText();
                PdfString richString = richText as PdfString;
                string text = richString.ToString();

                //check if richetext is formatted as black
                if (text.Contains("#000000"))
                {
                    int[] RGB = new int[] { 0, 0, 0 };
                    return RGB;
                }

                if (text.Contains("#FF"))
                {
                    string hexColor = text.Substring(text.LastIndexOf("#FF"), 7);
                    Color color = ColorTranslator.FromHtml(hexColor);
                    int[] RGB = new int[] { Convert.ToInt16(color.R), Convert.ToInt16(color.G), Convert.ToInt16(color.B) };
                    colors = RGB;
                    return colors;
                }

                /*
                //if not get color from default apperance
                if (daText.Length > 4)
                {
                    if (double.TryParse(daText[2], out var result2) && double.TryParse(daText[3], out var result3) && double.TryParse(daText[4], out var result4))
                    {
                        int[] RGB = new int[] { Convert.ToInt16(double.Parse(daText[2]) * 256), Convert.ToInt16(double.Parse(daText[3]) * 256), Convert.ToInt16(double.Parse(daText[4]) * 256) };
                        colors = RGB;
                        return colors;
                    }
                }

                //if not get color from default style
                if (colors == null && dsText[dsText.Length - 1].Contains("#"))
                {
                    string temp = dsText[dsText.Length - 1];
                    string hexColor = temp.Substring(temp.LastIndexOf(':') + 1);
                    Color color = ColorTranslator.FromHtml(hexColor);
                    int[] RGB = new int[] { Convert.ToInt16(color.R), Convert.ToInt16(color.G), Convert.ToInt16(color.B) };
                    colors = RGB;
                    return colors;
                }
                */
            }
            else if (annot.GetSubtype() == PdfName.PolyLine || annot.GetSubtype() == PdfName.Line || annot.GetSubtype() == PdfName.Polygon)
            {
                PdfArray tempColors = annot.GetColorObject();
                int[] RGB = new int[3];
                for (int j = 0; j < tempColors.Size(); j++)
                {
                    RGB[j] = Convert.ToInt16(double.Parse(tempColors.GetAsNumber(j).ToString()) * 256);
                }
                colors = RGB;
                return colors; ;
            }
            return colors;
        }
        private static bool checkIfColorIsSimilar(int[] baseColorRgb, int[] newColorRgb, int tolerance)
        {

            //RGB is a 3d space. each color is a point in the 255 domain.
            //calculate distance from the base color and check against tolerance

            int distance = ((newColorRgb[0] - baseColorRgb[0]) ^ 2 + (newColorRgb[1] - baseColorRgb[1]) ^ 2 + (newColorRgb[2] - baseColorRgb[2]) ^ 2) ^ (1 / 2);
            return (Math.Abs(distance) <= tolerance);
        }

        //####### update content

        public static void updateDrawingWithFcrAndRevision (string sourcePath, string tempPath, string fcrNumber, string revision)
        {
            //get pdf for read and write
            PdfDocument pdfDoc = getPdfForReadAndWrite(sourcePath, tempPath);
            //get page
            PdfPage page = pdfDoc.GetFirstPage();
            //get annotation on page
            IList<PdfAnnotation> annotations = page.GetAnnotations();

            //set regex 
            string fcrRegex = @"^[Ff][Cc][Rr]\s[0-9][0-9][0-9][0-9][0-9][0-9]\s*";
            string revisionRegex = @"^[A-Z]\.[0-9][0-9]\s*";

            for (int i = 0; i < annotations.Count; i++)
            {
                if (annotations[i].GetSubtype() == PdfName.FreeText)
                {
                    PdfAnnotation anno = annotations[i] as PdfAnnotation;
                    PdfMarkupAnnotation markupAnno = anno as PdfMarkupAnnotation;
                    string content = anno.GetContents().ToString();
                    if (content.Length > 2 && content.Length < 15)
                    {
                        if (Regex.IsMatch(content, fcrRegex))
                        {
                            string newText = String.Format(
                                "<?xml version=\"1.0\"?>" +
                                "<body xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:xfa=\"http://www.xfa.org/schema/xfa-data/1.0/\" xfa:APIVersion=\"Acrobat:11.0.0\" xfa:spec=\"2.0.2\">" +
                                "<p dir=\"ltr\">" +
                                "<span style=\"text-align:left;font-size:14pt;font-style:normal;font-weight:bold;color:#FF1418;font-family:Helvetica\">{0}" +
                                "</span>" +
                                "</p>" +
                                "</body>", "FCR " + fcrNumber);
                            PdfString newString = new PdfString(newText);
                            markupAnno.SetRichText(newString);
                        }
                        if (Regex.IsMatch(content, revisionRegex))
                        {
                            string newText = String.Format(
                                "<?xml version=\"1.0\"?>" +
                                "<body xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:xfa=\"http://www.xfa.org/schema/xfa-data/1.0/\" xfa:APIVersion=\"Acrobat:11.0.0\" xfa:spec=\"2.0.2\">" +
                                "<p dir=\"ltr\">" +
                                "<span style=\"text-align:left;font-size:14pt;font-style:normal;font-weight:bold;color:#FF1418;font-family:Helvetica\">{0}" +
                                "</span>" +
                                "</p>" +
                                "</body>", revision);
                            PdfString newString = new PdfString(newText);
                            markupAnno.SetRichText(newString);
                        }
                    }
                }
            }
            pdfDoc.Close();
        }
        public static void updateTextBoxesWithFcr(string sourcePath, string tempPath, string fcrNumber)
        {
            //get pdf for read and write
            PdfDocument pdfDoc = getPdfForReadAndWrite(sourcePath, tempPath);
            //get page
            PdfPage page = pdfDoc.GetFirstPage();
            //get annotation on page
            IList<PdfAnnotation> annotations = page.GetAnnotations();

            //set regex 
            string regexPattern = @"[Ff][Cc][Rr]\s[0-9][0-9][0-9][0-9][0-9][0-9]\s*";

            for (int i = 0; i < annotations.Count; i++)
            {
                if (annotations[i].GetSubtype() == PdfName.FreeText)
                {
                    PdfAnnotation anno = annotations[i] as PdfAnnotation;
                    PdfMarkupAnnotation markupAnno = anno as PdfMarkupAnnotation;
                    string content = anno.GetContents().ToString();
                    if (content.Length>4 && content.Length<15)
                    {
                        if (Regex.IsMatch(content, regexPattern))
                        {
                            string newText = String.Format(
                                "<?xml version=\"1.0\"?>" +
                                "<body xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:xfa=\"http://www.xfa.org/schema/xfa-data/1.0/\" xfa:APIVersion=\"Acrobat:11.0.0\" xfa:spec=\"2.0.2\">" +
                                "<p dir=\"ltr\">" +
                                "<span style=\"text-align:left;font-size:14pt;font-style:normal;font-weight:bold;color:#FF1418;font-family:Helvetica\">{0}" +
                                "</span>" +
                                "</p>" +
                                "</body>", "FCR " + fcrNumber);

                            PdfString newString = new PdfString(newText);
                            markupAnno.SetRichText(newString);
                        }
                    }
                }
            }
            pdfDoc.Close();
        }
        public static void updateTextBoxesWithRevision(string sourcePath, string tempPath, string revision)
        {
            //get pdf for read and write
            PdfDocument pdfDoc = getPdfForReadAndWrite(sourcePath, tempPath);
            //get page
            PdfPage page = pdfDoc.GetFirstPage();
            //get annotation on page
            IList<PdfAnnotation> annotations = page.GetAnnotations();

            //set regex 
            string regexPattern = @"[A-Z].[0-9][0-9]\s*";

            for (int i = 0; i < annotations.Count; i++)
            {
                if (annotations[i].GetSubtype() == PdfName.FreeText)
                {
                    PdfAnnotation anno = annotations[i] as PdfAnnotation;
                    PdfMarkupAnnotation markupAnno = anno as PdfMarkupAnnotation;
                    string content = anno.GetContents().ToString();
                    if (content.Length < 10)
                    {
                        if (Regex.IsMatch(content, regexPattern))
                        {
                            string newText = String.Format(
                                "<?xml version=\"1.0\"?>" +
                                "<body xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:xfa=\"http://www.xfa.org/schema/xfa-data/1.0/\" xfa:APIVersion=\"Acrobat:11.0.0\" xfa:spec=\"2.0.2\">" +
                                "<p dir=\"ltr\">" +
                                "<span style=\"text-align:left;font-size:14pt;font-style:normal;font-weight:bold;color:#FF1418;font-family:Helvetica\">{0}" +
                                "</span>" +
                                "</p>" +
                                "</body>", revision);
                            PdfString newString = new PdfString(newText);
                            markupAnno.SetRichText(newString);
                        }
                    }

                }
            }
            pdfDoc.Close();
        }

        //################################################## WORK IN PROGRESS


        public static string pdfText(string path)
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

        public static string pdfTextByRect(string path)
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

        public static void addText(string path)
        {
            PdfWriter writer = new PdfWriter(path);
            PdfDocument pdfDocument = new PdfDocument(writer);
            Document doc = new Document(pdfDocument);

            Paragraph paragraph = new Paragraph("test string");
            //paragraph.SetFixedPosition(500, 100, 100);
            doc.Add(paragraph);
            doc.Close();
        }

        public static int getAnotationsCount(string path)
        {
            PdfReader reader = new PdfReader(path);
            PdfDocument pdfDoc = new PdfDocument(reader);
            Document doc = new Document(pdfDoc);
            PdfPage page = pdfDoc.GetPage(1);
            IList<PdfAnnotation> annotations = page.GetAnnotations();
            return annotations.Count;
        }
    }
}
