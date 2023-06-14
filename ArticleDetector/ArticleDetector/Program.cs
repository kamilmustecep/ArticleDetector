﻿using System;
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
            Console.Write("Denetleme yapılacak dosya ismini yapıştırınız : ");
            string fileName = Console.ReadLine();
            Console.WriteLine();

            //string dosyaYolu = "C:\\Users\\Asus\\Desktop\\test.docx";
            string dosyaYolu = "C:\\Users\\Asus\\Desktop\\articles\\"+ fileName.Replace(".docx", "") + ".docx";


            bool girintiKontrol = CheckIndentationAfterHeading(dosyaYolu);
            bool boslukKontrol = KontrolEtKenarBosluk(dosyaYolu);
            bool fontKontrol = KontrolEtFont(dosyaYolu);
            bool puntoKontrol = KontrolEtPunto(dosyaYolu);
            bool satirAraligiKontrol = KontrolEtSatirAraligi(dosyaYolu);

            Console.Write("\n[?] Atıf kontrolü için ilk atıf ismini (makale yazarının) girin. (Örn. 'Çakmaklı' yada 'Muradov')  : ");
            string atıfFirstName = Console.ReadLine();

            var atiflar = FindCitationsInDocx(dosyaYolu, atıfFirstName);

            Console.ReadLine();
            
        }

        

        //OK
        static bool KontrolEtFont(string dosyaYolu)
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
                            if (!run.RunProperties.RunFonts.Ascii?.Value.Equals("Times New Roman") ?? true)
                            {
                                Console.WriteLine("FONT : X (\""+ run.RunProperties.RunFonts.Ascii?.Value + "\" bulundu)");
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
                Console.WriteLine("FONT : BULUNAMADI! (MANUEL KONTROL ET)");
            }
            else
            {
                Console.WriteLine("FONT : OK");
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
                        //int length = match.Value.Length;
                        //string value = match.Value;

                        //if (length >= 100)
                        //{
                        //    value = match.Value.Substring(length - 50, 50);
                        //}


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

                        if (Math.Abs(topMarginCm - desiredMarginCm) < 0.02 && Math.Abs(bottomMarginCm - desiredMarginCm) < 0.02 &&
                            Math.Abs(leftMarginCm - desiredMarginCm) < 0.02 && Math.Abs(rightMarginCm - desiredMarginCm) < 0.02)
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
                    .FirstOrDefault(p => p.InnerText=="Giriş" || p.InnerText.ToLower()=="giriş" || p.InnerText.ToLower().Trim().EndsWith("giriş") || p.InnerText.ToLower().Trim().EndsWith("introduction") || p.InnerText=="INTRODUCTION" || p.InnerText.Trim().EndsWith("Introduction"));


                if (startParagraph != null)
                {
                    // "Giriş" paragrafının sonraki paragrafları kontrol etme
                    bool hasIndentation = false;

                    foreach (Paragraph paragraph in startParagraph.ElementsAfter().Where(x=>x.GetType().Name== "Paragraph"))
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

                                if ((paragraph.ParagraphProperties?.Indentation?.FirstLine?.Value == "0" || paragraph.ParagraphProperties?.Indentation?.FirstLine?.Value == "null" || paragraph.ParagraphProperties?.Indentation?.FirstLine?.Value == null) && paragraph.ParagraphProperties?.NumberingProperties?.NumberingId==null)
                                {
                                    if (!paragraph.InnerText.StartsWith("       "))
                                    {
                                        Console.WriteLine("--- Aşağıdaki paragrafta GİRİNTİ YOK! ---");
                                        Console.WriteLine("- - - - - - - - - - - - - - - - - - - - -");
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
                                    Console.WriteLine("GİRİNTİ : OK");
                                }
                                else
                                {
                                    Console.WriteLine("GİRİNTİ : X");
                                }

                                return status;
                            }
                        }
                        
                    }
                }
                else
                {
                    Console.WriteLine("Giriş (Introduction) başlığı bulunamadı! Denetleme yapılamıyor.");
                }
            }

            if (status)
            {
                Console.WriteLine("GİRİNTİ : OK");
            }
            else
            {
                Console.WriteLine("GİRİNTİ : X");
            }

            Console.WriteLine("Kaynakça (References) başlığı bulunamadı! Denetleme son sayfaya kadar yapıldı.");

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


    }
}
