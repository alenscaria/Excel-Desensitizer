using System;
using System.IO;
using OfficeOpenXml;
using System.Linq;

class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string filePath = "<FILE_PATH>";
        string outputFilePath = "<OUTPUT_FILE_PATH>";

        try
        {

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[1];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;
                Console.WriteLine(worksheet + " " + rowCount + " " + colCount);

                Random random = new Random();
                DateTime startDate = new DateTime(2023, 1, 1);
                DateTime endDate = new DateTime(2025, 1, 1);
                int range = (endDate - startDate).Days;
                DateTime randomDate = startDate.AddDays(random.Next(range));

                Console.WriteLine("Starting updation of records...");
                for (int row = 2; row <= rowCount; row++)
                {
                    randomDate = startDate.AddDays(random.Next(range));
                    worksheet.Cells[row, 1].Value = randomDate.ToShortDateString();
                    worksheet.Cells[row, 2].Value = random.Next(80, 5000) + " Prop_Addr"; 
                    worksheet.Cells[row, 3].Value = "City_" + random.Next(1, 100); //replace address
                    worksheet.Cells[row, 4].Value = "User_" + row; //replace name
                    worksheet.Cells[row, 5].Value = "user" + row + "@example.com"; //replace email
                    worksheet.Cells[row, 6].Value = "999999" + random.Next(1000, 9999); //replace phone num
                    worksheet.Cells[row, 7].Value = random.Next(100000, 999999); //replace Ids
                }
                Console.WriteLine("Saving File...");
                package.SaveAs(new FileInfo(outputFilePath));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Exception:" + ex);
        }
        Console.WriteLine("Removed Sensitive Data successfully!\nSaved to " + outputFilePath);
    }
}
