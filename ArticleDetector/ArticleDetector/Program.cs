using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ArticleDetector
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string dosyaYolu = "C:\\Users\\Asus\\Desktop\\test.docx";

            bool girintiKontrol = KontrolEtGirinti(dosyaYolu);
            bool boslukKontrol = KontrolEtBosluk(dosyaYolu);
            bool fontKontrol = KontrolEtFont(dosyaYolu);
            bool puntoKontrol = KontrolEtPunto(dosyaYolu);
            bool satirAraligiKontrol = KontrolEtSatirAraligi(dosyaYolu);

            if (girintiKontrol && boslukKontrol && fontKontrol && puntoKontrol && satirAraligiKontrol)
            {
                Console.WriteLine("Makale uygun formatta.");
            }
            else
            {
                Console.WriteLine("Makale uygun formatta değil.");
            }
        }

        

        //OK
        static bool KontrolEtFont(string dosyaYolu)
        {
            using (WordprocessingDocument belge = WordprocessingDocument.Open(dosyaYolu, false))
            {
                Body body = belge.MainDocumentPart.Document.Body;

                foreach (Run run in body.Descendants<Run>())
                {
                    if (run.RunProperties != null && run.RunProperties.RunFonts != null)
                    {
                        if (!run.RunProperties.RunFonts.Ascii?.Value.Equals("Times New Roman") ?? true)
                        {
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        //OK
        static bool KontrolEtPunto(string dosyaYolu)
        {
            using (WordprocessingDocument belge = WordprocessingDocument.Open(dosyaYolu, false))
            {
                Body body = belge.MainDocumentPart.Document.Body;

                foreach (Run run in body.Descendants<Run>())
                {
                    if (run.RunProperties != null && run.RunProperties.FontSize != null)
                    {
                        if (!run.RunProperties.FontSize.Val?.Value.Equals("20") ?? true)
                        {
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        //OK
        static bool KontrolEtSatirAraligi(string dosyaYolu)
        {
            using (WordprocessingDocument belge = WordprocessingDocument.Open(dosyaYolu, false))
            {
                Body body = belge.MainDocumentPart.Document.Body;

                foreach (ParagraphProperties paragraphProperties in body.Descendants<ParagraphProperties>())
                {
                    if (paragraphProperties.SpacingBetweenLines != null)
                    {
                        if ((!paragraphProperties.SpacingBetweenLines.LineRule?.Value.Equals(LineSpacingRuleValues.Auto) ?? true)
                            || (!paragraphProperties.SpacingBetweenLines.Line?.Value.Equals("360") ?? true))
                        {
                            return false;
                        }
                    }
                }
            }

            return true;
        }




        static bool KontrolEtGirinti(string dosyaYolu)
        {
            using (WordprocessingDocument belge = WordprocessingDocument.Open(dosyaYolu, false))
            {
                Body body = belge.MainDocumentPart.Document.Body;

                foreach (Paragraph paragraph in body.Elements<Paragraph>())
                {
                    if (!string.IsNullOrWhiteSpace(paragraph.InnerText))
                    {
                        if (!paragraph.ParagraphProperties.Indentation?.Left?.Value.Equals(-720) ?? true)
                        {
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        static bool KontrolEtBosluk(string dosyaYolu)
        {
            using (WordprocessingDocument belge = WordprocessingDocument.Open(dosyaYolu, false))
            {
                Body body = belge.MainDocumentPart.Document.Body;

                foreach (Paragraph paragraph in body.Elements<Paragraph>())
                {
                    if (!string.IsNullOrWhiteSpace(paragraph.InnerText))
                    {
                        if ((!paragraph.ParagraphProperties.SpacingBetweenLines?.Before?.Value.Equals(360) ?? true)
                            || (!paragraph.ParagraphProperties.SpacingBetweenLines?.After?.Value.Equals(360) ?? true))
                        {
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        
    }
}
