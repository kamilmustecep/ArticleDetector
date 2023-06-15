using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;

namespace ArticleDetector
{
    internal class Program
    {

        public static List<string> names = new List<string>();
        public static List<string> shorts = new List<string>();
        public static List<string> kaynakcaNames = new List<string>();
        public static List<string> KaynakcaNotFound = new List<string>();

        static void Main(string[] args)
        {

            string jsonFilePath = "settings.txt";
            string jsonContent = File.ReadAllText(jsonFilePath);

            Models.settings settings = JsonConvert.DeserializeObject<Models.settings>(jsonContent);


            Console.Write("DENETLEME YAPILACAK DOSYA İSMİ : ");
            string fileName = Console.ReadLine();
            Console.WriteLine();

            string dosyaYolu = settings.filePath + fileName.Replace(".docx", "") + ".docx";
            //string dosyaYolu = "C:\\Users\\Asus\\Desktop\\articles\\"+ fileName.Replace(".docx", "") + ".docx";

            bool fontKontrol = KontrolEtFont(dosyaYolu,settings.fontFamily);
            bool puntoKontrol = KontrolEtPunto(dosyaYolu,settings.puntoPx);
            bool satirAraligiKontrol = KontrolEtSatirAraligi(dosyaYolu,settings.satirAraligi);
            bool boslukKontrol = KontrolEtKenarBosluk(dosyaYolu,settings.kenarBoslugu);
            bool girintiKontrol = CheckIndentationAfterHeading(dosyaYolu);
            

            Console.Write("\n[?] Atıf kontrolü için ilk atıf ismini (makale yazarının) girin. (Örn. 'Çakmaklı' yada 'Muradov')  : ");
            string atıfFirstName = Console.ReadLine();

            var atiflar = FindCitationsInDocx(dosyaYolu, atıfFirstName);

            var a = KaynakcaIsimKontrol(dosyaYolu);


            Console.ReadLine();

        }



        //OK
        static bool KontrolEtFont(string dosyaYolu,string fontFamily)
        {
            bool haveInformation = false;

            using (WordprocessingDocument belge = WordprocessingDocument.Open(dosyaYolu, false))
            {
                Body body = belge.MainDocumentPart.Document.Body;

                foreach (Run run in body.Descendants<Run>())
                {
                    if (run.RunProperties != null && run.RunProperties.RunFonts != null)
                    {
                        if (run.RunProperties.RunFonts.Ascii != null)
                        {
                            if (!run.RunProperties.RunFonts.Ascii?.Value.Equals(fontFamily) ?? true)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("@     FONT : X (\"" + run.RunProperties.RunFonts.Ascii?.Value + "\" bulundu)");
                                Console.ForegroundColor = ConsoleColor.White;
                                return false;
                            }
                            else
                            {
                                haveInformation = true;
                            }
                        }
                    }
                }
            }

            if (!haveInformation)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("@     FONT : BULUNAMADI! (MANUEL KONTROL ET)");
                Console.ForegroundColor = ConsoleColor.White;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("@     FONT : OK");
                Console.ForegroundColor = ConsoleColor.White;
            }



            return true;
        }

        //OK
        static bool KontrolEtPunto(string dosyaYolu, string puntoPx)
        {
            using (WordprocessingDocument belge = WordprocessingDocument.Open(dosyaYolu, false))
            {
                Body body = belge.MainDocumentPart.Document.Body;

                foreach (Run run in body.Descendants<Run>())
                {
                    if (run.RunProperties != null && run.RunProperties.FontSize != null)
                    {
                        int puntopxInt = Convert.ToInt32(puntoPx)*2;

                        if (!run.RunProperties.FontSize.Val?.Value.Equals(puntopxInt.ToString()) ?? true)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("@     PUNTO : X");
                            Console.ForegroundColor = ConsoleColor.White;
                            return false;
                        }
                    }
                }
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("@     PUNTO : OK");
            Console.ForegroundColor = ConsoleColor.White;
            return true;
        }

        //OK - "satirAraligi" parameter do not using
        static bool KontrolEtSatirAraligi(string dosyaYolu, string satirAraligi)
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
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("@     SATIR ARALIĞI : X");
                            Console.ForegroundColor = ConsoleColor.White;
                            return false;
                        }
                    }
                }
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("@     SATIR ARALIĞI : OK");
            Console.ForegroundColor = ConsoleColor.White;
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

                    //string pattern1 = @"(?<!\()\bÇakmaklı\b[^)]*\)";
                    //string pattern2 = @"\([^)]*Çakmaklı[^)]*\)";
                    string pattern = $@"(?<!\()\b{atifFirstName}\b[^)]*\)|\([^)]*{atifFirstName}[^)]*\)";

                    MatchCollection matches1 = Regex.Matches(paragraphText, pattern);

                    foreach (Match match in matches1)
                    {
                        citations.Add(match.Value);
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

                    if (paragraph.InnerText.ToLower() == "kaynakça" || paragraph.InnerText.ToLower() == "references" || paragraph.InnerText.ToLower().Trim().EndsWith("references") || paragraph.InnerText.ToLower().Trim().EndsWith("kaynaklar"))
                    {
                        break;
                    }


                }

                int sayac = 1;

                Console.WriteLine("\n[ - - - - - BELGEDE GEÇEN ATIFLAR - - - - - ]\n");
                Console.ForegroundColor = ConsoleColor.Yellow;
                foreach (var atif in citations)
                {
                    Console.WriteLine(sayac + ". " + atif);
                    sayac++;
                }
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write("\n[?] KAYNAKÇA'DA YER ALMAYAN İSİMLERİ GÖRMEK İÇİN 'E veya e' yazın : ");
                detailRequest = Console.ReadLine();

                if (detailRequest == "E" || detailRequest == "e")
                {
                    sayac = 1;
                    //Console.WriteLine("\n[ - - - - - -BELGEDE GEÇEN TÜM ATIF ve KISALTMALAR - - - - - ] \n");

                    foreach (var atif in allcitations)
                    {

                        List<string> ozelIsimler = AyiklaOzelIsimler(atif);
                        bool sayiVarmi = VeriTurunuKontrolEt(atif);

                        if (ozelIsimler != null && ozelIsimler.Count > 0)
                        {
                            if (sayiVarmi)
                            {
                                foreach (var isim in ozelIsimler)
                                {
                                    if (!names.Any(x => x == isim))
                                    {
                                        names.Add(isim);
                                    }

                                }
                            }
                            else
                            {
                                foreach (var shortName in ozelIsimler)
                                {
                                    if (!shorts.Any(x => x == shortName) && !names.Any(x => x == shortName))
                                    {
                                        shorts.Add(shortName);
                                    }
                                }
                            }

                        }
                        //Console.WriteLine(sayac + ". " + atif);
                        sayac++;
                    }
                }
            }

            return citations;
        }

        //OK
        static bool KontrolEtKenarBosluk(string dosyaYolu, string kenarBoslugu)
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

                        double desiredMarginCm = Convert.ToDouble(kenarBoslugu);

                        if (Math.Abs(topMarginCm - desiredMarginCm) < 0.02 && Math.Abs(bottomMarginCm - desiredMarginCm) < 0.02 &&
                            Math.Abs(leftMarginCm - desiredMarginCm) < 0.02 && Math.Abs(rightMarginCm - desiredMarginCm) < 0.02)
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("@     KENAR BOŞLUKLARI : OK\n");
                            Console.ForegroundColor = ConsoleColor.White;
                            return true;
                        }
                    }
                }
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("@     KENAR BOŞLUKLARI : X");
                Console.ForegroundColor = ConsoleColor.White;
                return false;
            }

        }


        static bool CheckIndentationAfterHeading(string filePath)
        {
            bool status = true;
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                DocumentFormat.OpenXml.Wordprocessing.Document document = wordDoc.MainDocumentPart.Document;
                Body body = document.Body;

                // "Giriş" kelimesini içeren paragrafı bulma
                List<Paragraph> startParagraphs = body.Descendants<Paragraph>().ToList();


                Paragraph startParagraph = body.Descendants<Paragraph>()
                    .FirstOrDefault(p => p.InnerText == "Giriş" || p.InnerText.ToLower() == "giriş" || p.InnerText.ToLower().Trim().EndsWith("giriş") || p.InnerText.ToLower().Trim().EndsWith("introduction") || p.InnerText == "INTRODUCTION" || p.InnerText.Trim().EndsWith("Introduction"));

            x:

                if (startParagraph != null)
                {
                    // "Giriş" paragrafının sonraki paragrafları kontrol etme

                    foreach (Paragraph paragraph in startParagraph.ElementsAfter().Where(x => x.GetType().Name == "Paragraph"))
                    {

                        if (!String.IsNullOrEmpty(paragraph.InnerText))
                        {
                            bool isTitle = false;
                            //bool isMadde = false;
                            foreach (Run run in paragraph.Descendants<Run>())
                            {
                                if (run.RunProperties?.Bold != null)
                                {
                                    isTitle = true;
                                }
                            }


                            // Paragrafın girintisi varsa hasIndentation değerini true yap ve döngüden çık
                            if (!isTitle && paragraph.InnerText.Length >= 150)
                            {
                                ParagraphStyleId styleId = paragraph.ParagraphProperties?.ParagraphStyleId;

                                if ((paragraph.ParagraphProperties?.Indentation?.FirstLine?.Value == "0" || paragraph.ParagraphProperties?.Indentation?.FirstLine?.Value == "null" || paragraph.ParagraphProperties?.Indentation?.FirstLine?.Value == null) && paragraph.ParagraphProperties?.NumberingProperties?.NumberingId == null)
                                {
                                    if (!paragraph.InnerText.StartsWith("       "))
                                    {
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine("--- Aşağıdaki paragrafta GİRİNTİ YOK! ---");
                                        Console.WriteLine("- - - - - - - - - - - - - - - - - - - - -");
                                        Console.ForegroundColor = ConsoleColor.White;
                                        Console.WriteLine(paragraph.InnerText + "\n");
                                        status = false;
                                    }

                                }
                                else
                                {
                                    //Console.WriteLine(" Girinti VAR ");
                                    //Console.WriteLine(paragraph.InnerText + "\n");/R
                                }
                            }

                            if (paragraph.InnerText.ToLower() == "kaynakça" || paragraph.InnerText.ToLower() == "references" || paragraph.InnerText.ToLower().Trim().EndsWith("references") || paragraph.InnerText.ToLower().Trim().EndsWith("kaynaklar"))
                            {
                                if (status)
                                {
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("@     GİRİNTİ : OK");
                                    Console.ForegroundColor = ConsoleColor.White;
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("@     GİRİNTİ : X");
                                    Console.ForegroundColor = ConsoleColor.White;
                                }

                                return status;
                            }
                        }

                    }
                }
                else
                {
                    Console.WriteLine("Giriş (Introduction) başlığı bulunamadı! Denetleme yapılamıyor.");

                    Console.Write("\n[?] Giriş Başlık metnini manuel yazarak arama yap : ");
                    string girisHead = Console.ReadLine();

                    startParagraph = body.Descendants<Paragraph>()
                    .FirstOrDefault(p => p.InnerText.ToLower() == girisHead.ToLower() || p.InnerText.ToLower() == girisHead.ToLower() || p.InnerText.ToLower().Trim().EndsWith(girisHead.ToLower()) || p.InnerText.ToLower().Trim().EndsWith(girisHead.ToLower()));

                    goto x;
                }
            }

            if (status)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("GİRİNTİ : OK");
                Console.ForegroundColor = ConsoleColor.White;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("GİRİNTİ : X");
                Console.ForegroundColor = ConsoleColor.White;
            }

            Console.WriteLine("Kaynakça (References) başlığı bulunamadı! Denetleme son sayfaya kadar yapıldı.");

            return status;
        }


        static bool KaynakcaIsimKontrol(string filePath)
        {
            bool status = true;
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                DocumentFormat.OpenXml.Wordprocessing.Document document = wordDoc.MainDocumentPart.Document;
                Body body = document.Body;

                // "Giriş" kelimesini içeren paragrafı bulma
                List<Paragraph> startParagraphs = body.Descendants<Paragraph>().ToList();

                string ResourcefullTextSearch = "";

                string ArticlefullTextSearch = "";

                Paragraph startParagraph = body.Descendants<Paragraph>()
                    .FirstOrDefault(p => p.InnerText.ToLower() == "kaynakça" || p.InnerText.ToLower() == "references" || p.InnerText.ToLower().Trim().EndsWith("references") || p.InnerText.ToLower().Trim().EndsWith("kaynaklar"));

            x:

                if (startParagraph != null)
                {


                    foreach (Paragraph paragraph in startParagraph.ElementsAfter().Where(x => x.GetType().Name == "Paragraph"))
                    {

                        if (!String.IsNullOrEmpty(paragraph.InnerText))
                        {
                            ResourcefullTextSearch = ResourcefullTextSearch + " " + paragraph.InnerText;
                        }

                    }



                    foreach (var name in names)
                    {
                        if (!ResourcefullTextSearch.Contains(name))
                        {
                            if (!ResourcefullTextSearch.Contains(name.Replace(" ", "")))
                            {
                                KaynakcaNotFound.Add(name);
                            }
                        }
                    }

                    Console.WriteLine("\n[ - - - - - - KAYNAKÇA'DA YER ALMAYAN KELİMELER  - - - - - ] \n");

                    Console.ForegroundColor = ConsoleColor.Yellow;
                    int sayac2 = 1;
                    foreach (var nameExtract in KaynakcaNotFound)
                    {
                        Console.WriteLine(sayac2 + ". " + nameExtract);
                        sayac2++;
                    }
                    Console.ForegroundColor = ConsoleColor.White;
                }
                else
                {
                    Console.WriteLine("Kaynakça (resources) başlığı bulunamadı! Denetleme yapılamıyor.");

                    Console.Write("\n[?] Kaynakça Başlık metnini manuel yazarak arama yap : ");
                    string kaynakcaHead = Console.ReadLine();

                    startParagraph = body.Descendants<Paragraph>()
                    .FirstOrDefault(p => p.InnerText.ToLower() == kaynakcaHead.ToLower() || p.InnerText.ToLower() == kaynakcaHead.ToLower() || p.InnerText.ToLower().Trim().EndsWith(kaynakcaHead.ToLower()) || p.InnerText.ToLower().Trim().EndsWith(kaynakcaHead.ToLower()));

                    goto x;

                }



                startParagraph = body.Descendants<Paragraph>()
            .FirstOrDefault(p => p.InnerText == "Giriş" || p.InnerText.ToLower() == "giriş" || p.InnerText.ToLower().Trim().EndsWith("giriş") || p.InnerText.ToLower().Trim().EndsWith("introduction") || p.InnerText == "INTRODUCTION" || p.InnerText.Trim().EndsWith("Introduction"));

            y:

                if (startParagraph != null)
                {
                    // "Giriş" paragrafının sonraki paragrafları kontrol etme

                    foreach (Paragraph paragraph in startParagraph.ElementsAfter().Where(x => x.GetType().Name == "Paragraph"))
                    {

                        if (!String.IsNullOrEmpty(paragraph.InnerText))
                        {
                            ArticlefullTextSearch = ArticlefullTextSearch + " " + paragraph.InnerText;


                            if (paragraph.InnerText.ToLower() == "kaynakça" || paragraph.InnerText.ToLower() == "references" || paragraph.InnerText.ToLower().Trim().EndsWith("references") || paragraph.InnerText.ToLower().Trim().EndsWith("kaynaklar"))
                            {
                                break;
                            }
                        }

                    }
                }
                else
                {
                    Console.WriteLine("Giriş (Introduction) başlığı bulunamadı! Denetleme yapılamıyor.");

                    Console.Write("\n[?] Giriş Başlık metnini manuel yazarak arama yap : ");
                    string girisHead = Console.ReadLine();

                    startParagraph = body.Descendants<Paragraph>()
                    .FirstOrDefault(p => p.InnerText.ToLower() == girisHead.ToLower() || p.InnerText.ToLower() == girisHead.ToLower() || p.InnerText.ToLower().Trim().EndsWith(girisHead.ToLower()) || p.InnerText.ToLower().Trim().EndsWith(girisHead.ToLower()));

                    goto y;
                }




                List<string> kaynakcaParantezler = new List<string>();

                string otherPattern = @"(?<!\()\w+\b[^)]*\)|\([^)]*\w+\b[^)]*\)";

                MatchCollection matches2 = Regex.Matches(ResourcefullTextSearch, otherPattern);

                foreach (Match match in matches2)
                {
                    int length = match.Value.Length;
                    string value = match.Value;

                    if (length >= 100)
                    {
                        value = match.Value.Substring(length - 50, 50);
                    }
                    kaynakcaParantezler.Add(value);
                }


                //Kaynakça'daki özel isimler

                foreach (var kparantezler in kaynakcaParantezler)
                {


                    List<string> ozelIsimler = AyiklaOzelIsimler(kparantezler);

                    if (ozelIsimler != null && ozelIsimler.Count > 0)
                    {

                        foreach (var isim in ozelIsimler)
                        {
                            if (!kaynakcaNames.Any(x => x == isim) && !names.Any(x => x == isim) && !shorts.Any(x => x == isim))
                            {
                                if (!ArticlefullTextSearch.Contains(isim))
                                {
                                    kaynakcaNames.Add(isim);
                                }
                                else
                                {
                                    var asda = "var";
                                }
                                
                            }

                        }
                    }

                }

                Console.WriteLine("\n[ - - - - - - MAKALE'DE YER ALMAYAN KELİMELER  - - - - - ] \n");

                Console.ForegroundColor = ConsoleColor.Yellow;
                int sayac = 1;
                foreach (var nameExtract in kaynakcaNames)
                {
                    Console.WriteLine(sayac + ". " + nameExtract);
                    sayac++;
                }
                Console.ForegroundColor = ConsoleColor.White;
            }

            return status;
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

        static List<string> AyiklaOzelIsimler(string metin)
        {
            List<string> ozelIsimler = new List<string>();

            TextInfo textInfo = new CultureInfo("tr-TR", false).TextInfo;
            string duzenliIfadeDeseni = @"\b[A-ZÇĞİÖŞÜ][a-zçğıöşü]+(\s[A-ZÇĞİÖŞÜ][a-zçğıöşü]+)*\b";

            foreach (Match eslesme in Regex.Matches(metin, duzenliIfadeDeseni))
            {
                string isim = textInfo.ToTitleCase(eslesme.Value.ToLower());
                ozelIsimler.Add(isim);
            }

            return ozelIsimler;
        }


        static bool VeriTurunuKontrolEt(string metin)
        {
            string duzenliIfadeDeseni = @"\((\d+)\)";
            Match eslesme = Regex.Match(metin, duzenliIfadeDeseni);

            if (eslesme.Success)
            {
                return true; // Parantez içinde sayı var
            }
            else
            {
                return false; // Parantez içinde metin var
            }
        }
    }
}
