using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection.PortableExecutable;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using OfficeOpenXml;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

public class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        while (true)
        {
            var folderPath = "";
            var dialog = new OpenFileDialog
            {
                Multiselect = false,
                Title = "Pdf Belgesi Seçiniz",
                Filter = "Pdf Document|*.pdf"
            };
            var excelName = "";
            using (dialog)
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    folderPath = dialog.FileName;
                    excelName = dialog.SafeFileName;
                    excelName = excelName.Substring(0,excelName.IndexOf("."));
                }
            }

            Dictionary<string, int> uniqueWords = new Dictionary<string, int>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Unique Words");

            var filePath = folderPath;
            var reader = PdfDocument.Open(filePath);


            foreach (var page in reader.GetPages())
            {
                var pageText = ContentOrderTextExtractor.GetText(page);
                pageText = pageText.Replace("\n", " ").ToLower();
                pageText = pageText.Replace(":", "").ToLower();
                var TabLength = 4;
                var TabSpace = new String(' ', TabLength);
                pageText = pageText.Replace("\t", TabSpace);
                RegexOptions options = RegexOptions.None;
                Regex rgx = new Regex("[^a-zA-ZğüşöçıIİĞÜŞÖÇ - ]");
                Regex regex = new Regex("[ ]{2,}", options);
                pageText = rgx.Replace(pageText, "");
                pageText = regex.Replace(pageText, " ");
                string[] words = pageText.Split(' ');
                foreach (string word in words)
                {
                    if (uniqueWords.ContainsKey(word))
                    {
                        uniqueWords[word]++;
                    }
                    else
                    {
                        uniqueWords.Add(word, 1);
                    }
                }

            }
            int row = 1;
            foreach (var word in uniqueWords)
            {
                worksheet.Cells[row, 1].Value = word.Key;
                worksheet.Cells[row, 2].Value = word.Value;
                row++;
                Console.WriteLine($"{word.Key}: {word.Value}");
            }
            if (!Directory.Exists(@"C:\CountWords\"))
            {
                Directory.CreateDirectory(@"C:\CountWords\");
            }
            if (File.Exists(@"C:\CountWords\" + excelName + ".xlsx"))
            {
                File.Delete(@"C:\CountWords\" + excelName + ".xlsx");
            }
            excelPackage.SaveAs(@"C:\CountWords\" + excelName + ".xlsx");
            Console.WriteLine("Devam etmek için Space(Boşluk) tuşuna basınız! Çıkmak için iki kere Enter tuşuna basınız");
            ConsoleKeyInfo keyPressed = Console.ReadKey();
            if (keyPressed.Key != ConsoleKey.Spacebar)
            {
                break;
            }
        }

    }
}