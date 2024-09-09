using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        // Array of DOCX file paths
        string[] docxPaths = new string[]
        {
            @"C:\Users\YourName\Documents\file1.docx", // Change these paths to your input DOCX files
            @"C:\Users\YourName\Documents\file2.docx",
            @"C:\Users\YourName\Documents\file3.docx"
        };

        ConvertDocxToExcel(docxPaths);
    }

    static void ConvertDocxToExcel(string[] docxPaths)
    {
        // Loop through each DOCX file
        foreach (var docxPath in docxPaths)
        {
            // Load the DOCX file
            using (WordprocessingDocument? wordDoc = WordprocessingDocument.Open(docxPath, false))
            {
                if (wordDoc == null)
                {
                    Console.WriteLine($"Failed to open the DOCX file: {docxPath}");
                    continue; // Skip this file if it cannot be opened
                }

                var body = wordDoc.MainDocumentPart?.Document.Body;
                if (body == null)
                {
                    Console.WriteLine($"The document does not contain any content: {docxPath}");
                    continue; // Skip if the document body is null
                }

                var paragraphs = body.Elements<Paragraph>().Select(p => p.InnerText).ToList();

                // Generate the corresponding output Excel file name
                string excelFileName = Path.GetFileNameWithoutExtension(docxPath) + ".xlsx";
                string excelPath = Path.Combine(Path.GetDirectoryName(docxPath), excelFileName);

                // Create a new Excel package
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    // Add a new worksheet for the document
                    var worksheet = excelPackage.Workbook.Worksheets.Add("Content");

                    // Write paragraphs to the Excel file
                    for (int i = 0; i < paragraphs.Count; i++)
                    {
                        worksheet.Cells[i + 1, 1].Value = paragraphs[i]; // Writing to column A
                    }

                    // Save the Excel file
                    FileInfo excelFile = new FileInfo(excelPath);
                    excelPackage.SaveAs(excelFile);
                }
            }
        }

        Console.WriteLine("Conversion completed for all documents!");
    }
}
