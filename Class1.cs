﻿using System;
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




namespace PdfParser
{
    public static class PdfParser
    {
        public static void copyAnnotsByColor(string sourcePath, string destPath)
        {
            //PdfDocument sourceDoc = getPdfForRead(sourcePath);
            PdfDocument destDoc = getPdfForReadAndWrite(sourcePath,destPath);

            //get annotations from source
            PdfPage sourcePage = destDoc.GetFirstPage();
            IList<PdfAnnotation> annotations = sourcePage.GetAnnotations();

            //filter annotations by color
            int[] redRGB = new int[] { 255, 0, 0 };
            excludeAnnotationsByColor(ref annotations,redRGB);

            //remove annotations in destination pdf
            PdfPage destPage = destDoc.GetFirstPage();
            foreach(PdfAnnotation annot in annotations)
            {
                destPage.RemoveAnnotation(annot);
            }

            //close documents
            //sourceDoc.Close();
            destDoc.Close();
        }
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
        private static bool checkIfColorIsSimilar(int[] baseColorRgb, int[] newColorRgb, int tolerance)
        {

            //RGB is a 3d space. each color is a point in the 255 domain.
            //calculate distance from the base color and check against tolerance

            int distance = ((newColorRgb[0] - baseColorRgb[0]) ^ 2 + (newColorRgb[1] - baseColorRgb[1]) ^ 2 + (newColorRgb[2] - baseColorRgb[2]) ^ 2) ^ (1 / 2);
            return (Math.Abs(distance) <= tolerance);
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
            }
            else if (annot.GetSubtype() == PdfName.PolyLine)
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
            else if (annot.GetSubtype() == PdfName.Line)
            {
                PdfArray tempColors = annot.GetColorObject();
                int[] RGB = new int[3];
                for (int j = 0; j < tempColors.Size(); j++)
                {
                    RGB[j] = Convert.ToInt16(double.Parse(tempColors.GetAsNumber(j).ToString()) * 256);
                }
                colors = RGB;
                return colors;
            }
            return colors;
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

        public static void updateTextBoxesWithFcr(string sourcePath, string destPath, string fcrNumber)
        {
            PdfReader reader = new PdfReader(sourcePath);
            PdfWriter writer = new PdfWriter(destPath);

            PdfDocument pdfDocSource = new PdfDocument(reader, writer);

            PdfPage sourcePage = pdfDocSource.GetFirstPage();

            IList<PdfAnnotation> annotations = sourcePage.GetAnnotations();
            
            for (int i = 0; i < annotations.Count; i++)
            {
                if (annotations[i].GetSubtype() == PdfName.FreeText)
                {
                    PdfAnnotation anno = annotations[i] as PdfAnnotation;
                    PdfMarkupAnnotation markupAnno = anno as PdfMarkupAnnotation;
                    PdfObject richText = markupAnno.GetRichText();
                    PdfString richString = richText as PdfString;
                    string text = richString.ToString();
                    int fcrIndex = -1;
                    if (text.Contains("FCR "))
                    {
                        fcrIndex = text.IndexOf("FCR ") + 4;
                    }
                    else if (text.Contains("FCR"))
                    {
                        fcrIndex = text.IndexOf("FCR") + 3;
                    }

                    if (fcrIndex > -1)
                    {
                        string newText = text.Substring(0, fcrIndex) + fcrNumber + text.Substring(fcrIndex + fcrNumber.Length, text.Length - (fcrIndex + fcrNumber.Length));
                        PdfString newString = new PdfString(newText);
                        markupAnno.SetRichText(newString);
                    }
                }
            }
            pdfDocSource.Close();
        }
    }
}