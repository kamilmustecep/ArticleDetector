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
            //string dosyaYolu = "C:\\Users\\Asus\\Desktop\\test.docx";
            string dosyaYolu = "C:\\Users\\kamil.mustecep\\Desktop\\test.docx";


            bool girintiKontrol = KontrolEtGirinti(dosyaYolu);
            bool boslukKontrol = KontrolEtKenarBosluk(dosyaYolu);
            bool fontKontrol = KontrolEtFont(dosyaYolu);
            bool puntoKontrol = KontrolEtPunto(dosyaYolu);
            bool satirAraligiKontrol = KontrolEtSatirAraligi(dosyaYolu);

            Console.WriteLine("\n[?] Atıf kontrolü için ilk atıf ismini girin. (Örn. 'Çakmaklı')  : ");
            string atıfFirstName = Console.ReadLine();

            var atiflar = FindCitationsInDocx(dosyaYolu, atıfFirstName);

            Console.ReadLine();

            
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
                            Console.WriteLine("FONT : X");
                            return false;
                        }
                    }
                }
            }

            Console.WriteLine("FONT : OK");

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
                            Console.WriteLine("PUNTO : X");
                            return false;
                        }
                    }
                }
            }
            Console.WriteLine("PUNTO : OK");
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
                            Console.WriteLine("SATIR ARALIĞI : X");
                            return false;
                        }
                    }
                }
            }

            Console.WriteLine("SATIR ARALIĞI : OK");
            return true;
        }

        //OK
        public static List<string> FindCitationsInDocx(string filePath, string atifFirstName)
        {
            List<string> citations = new List<string>();

            List<string> allcitations = new List<string>();

            string detailRequest = "";

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, false))
            {
                // Belgedeki tüm paragrafları al
                IEnumerable<Paragraph> paragraphs = document.MainDocumentPart.Document.Descendants<Paragraph>();

                foreach (Paragraph paragraph in paragraphs)
                {
                    string paragraphText = paragraph.InnerText;

                    // Paragraf metnindeki atıfları bulmak için kendi mantığınızı uygulayın
                    // Örneğin, belirli bir düzenli ifade veya kelime desenini arayabilirsiniz

                    // Örnek: Paragraf metninde parantez içindeki dört haneli sayıları bulma


                    //string pattern1 = @"(?<!\()\bÇakmaklı\b[^)]*\)";
                    //string pattern2 = @"\([^)]*Çakmaklı[^)]*\)";
                    string pattern = $@"(?<!\()\b{atifFirstName}\b[^)]*\)|\([^)]*{atifFirstName}[^)]*\)";

                    MatchCollection matches1 = Regex.Matches(paragraphText, pattern);

                    foreach (Match match in matches1)
                    {
                        int length = match.Value.Length;
                        string value = match.Value;

                        if (length >= 100)
                        {
                            value = match.Value.Substring(length - 50, 50);
                        }


                        citations.Add(value);
                    }


                    string otherPattern = @"(?<!\()\w+\b[^)]*\)|\([^)]*\w+\b[^)]*\)";

                    MatchCollection matches2 = Regex.Matches(paragraphText, otherPattern);

                    foreach (Match match in matches2)
                    {
                        int length = match.Value.Length;
                        string value = match.Value;

                        if (length >= 100)
                        {
                            value = match.Value.Substring(length - 50, 50);
                        }
                        allcitations.Add(value);
                    }




                }

                int sayac = 1;

                Console.WriteLine("\n[ - - - - - BELGEDE GEÇEN ATIFLAR - - - - - ]\n");
                foreach (var atif in citations)
                {
                    Console.WriteLine(sayac + ". " + atif);
                    sayac++;
                }

                Console.WriteLine("\n[?] KULLANILAN TÜM ATIF ve KISALTMALAR İÇİN 'E veya e' yazın : ");
                detailRequest = Console.ReadLine();

                if (detailRequest == "E" || detailRequest == "e")
                {
                    sayac = 1;
                    Console.WriteLine("\n[ - - - - - -BELGEDE GEÇEN TÜM ATIF ve KISALTMALAR - - - - - ] \n");
                    foreach (var atif in allcitations)
                    {
                        Console.WriteLine(sayac + ". " + atif);
                        sayac++;
                    }
                }
            }

            return citations;
        }

        //OK
        static bool KontrolEtKenarBosluk(string dosyaYolu)
        {
            using (WordprocessingDocument belge = WordprocessingDocument.Open(dosyaYolu, false))
            {
                // Word belgesinin ilk sayfasını al
                Body body = belge.MainDocumentPart.Document.Body;
                SectionProperties sectionProps = body.Elements<SectionProperties>().FirstOrDefault();

                if (sectionProps != null)
                {
                    // Kenar boşluklarını al
                    PageMargin pageMargin = sectionProps.Elements<PageMargin>().FirstOrDefault();

                    if (pageMargin != null)
                    {
                        // Kenar boşluklarını cm cinsinden kontrol et
                        double topMarginCm = ConvertToCm(pageMargin.Top.Value.ToString());
                        double bottomMarginCm = ConvertToCm(pageMargin.Bottom.Value.ToString());
                        double leftMarginCm = ConvertToCm(pageMargin.Left.Value.ToString());
                        double rightMarginCm = ConvertToCm(pageMargin.Right.Value.ToString());

                        double desiredMarginCm = 2.5;

                        if (Math.Abs(topMarginCm - desiredMarginCm) < 0.01 && Math.Abs(bottomMarginCm - desiredMarginCm) < 0.01 &&
                            Math.Abs(leftMarginCm - desiredMarginCm) < 0.01 && Math.Abs(rightMarginCm - desiredMarginCm) < 0.01)
                        {
                            Console.WriteLine("KENAR BOŞLUKLARI : OK");
                            return true;
                        }
                    }
                }

                Console.WriteLine("KENAR BOŞLUKLARI : X");
                return false;
            }

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
                            Console.WriteLine("GİRİNTİ : X");
                            return false;
                        }
                    }
                }
            }
            Console.WriteLine("GİRİNTİ : OK");
            return true;
        }

        

        //Other Functions

        static double ConvertToCm(string value)
        {
            double points = Convert.ToDouble(value);
            return points / 1440.0 * 2.54;
        }


    }
}
