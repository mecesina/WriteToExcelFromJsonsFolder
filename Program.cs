using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.IO;

namespace RethrieveJsonPropertiesFromFolder
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string folderPath = "Replace with the actual folder path";

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("EntityCodes");

                worksheet.Cells[1, 1].Value = "File Name";
                worksheet.Cells[1, 2].Value = "Entity Code";

                int row = 2;

                string[] jsonFiles = Directory.GetFiles(folderPath, "*.json", SearchOption.TopDirectoryOnly);

                foreach (string jsonFile in jsonFiles)
                {
                    string fileName = Path.GetFileName(jsonFile); //file name shall be generated as unique

                    string jsonContent = File.ReadAllText(jsonFile);

                    dynamic jsonObject = JsonConvert.DeserializeObject(jsonContent);
                    string entityCode = jsonObject.Product.EntityCode;

                    worksheet.Cells[row, 1].Value = fileName;
                    worksheet.Cells[row, 2].Value = entityCode;

                    row++;
                }

                string excelFilePath = "Replace with the desired file path.xlsx";
                excelPackage.SaveAs(new FileInfo(excelFilePath));
            }

            Console.WriteLine("Done! Entity codes extracted and saved to the Excel file.");
        }
    }
}
