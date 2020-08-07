using System;

namespace test
{
    class Program
    {

        //Create a new ExcelPackage
        using (ExcelPackage excelPackage = new ExcelPackage())
        {
           //Set some properties of the Excel document
           excelPackage.Workbook.Properties.Author = "VDWWD";
           excelPackage.Workbook.Properties.Title = "Title of Document";
           excelPackage.Workbook.Properties.Subject = "EPPlus demo export data";
           excelPackage.Workbook.Properties.Created = DateTime.Now;
            //Create the WorkSheet
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
            //Add some text to cell A1
            worksheet.Cells["A1"].Value = "My first EPPlus spreadsheet!";
            //You could also use [line, column] notation:
            worksheet.Cells[1, 2].Value = "This is cell B1!";
            //Save your file
            FileInfo fi = new FileInfo(@"Path\To\Your\File.xlsx");
            excelPackage.SaveAs(fi);
        }
    }
}
