using System;
using System.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace SampleWebApp.Core.Controllers
{
    public class HomeController : Controller
    {
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private readonly IHostingEnvironment _hostingEnvironment;

        public HomeController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        /// <summary>
        /// /Home/FileReport
        /// </summary>
        public IActionResult FileReport()
        {
            var fileDownloadName = "report.xlsx";
            var reportsFolder = "reports";

            using (var package = createExcelPackage())
            {
                package.SaveAs(new FileInfo(Path.Combine(_hostingEnvironment.WebRootPath, reportsFolder, fileDownloadName)));
            }
            return File($"~/{reportsFolder}/{fileDownloadName}", XlsxContentType, fileDownloadName);
        }

        public IActionResult Index()
        {
            return View();
        }


        public async Task<IActionResult> FileUpload(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return RedirectToAction("Index");
            }

            using (var memoryStream = new MemoryStream())
            {
                await file.CopyToAsync(memoryStream).ConfigureAwait(false);

                using (var package = new ExcelPackage(memoryStream))
                {
                    var worksheet = package.Workbook.Worksheets[1]; // Tip: To access the first worksheet, try index 1, not 0
                    return Content(readExcelPackageToString(package, worksheet));
                }
            }
        }


        private string readExcelPackage(FileInfo fileInfo, string worksheetName)
        {
            using (var package = new ExcelPackage(fileInfo))
            {
                return readExcelPackageToString(package, package.Workbook.Worksheets[worksheetName]);
            }
        }

        private string readExcelPackageToString(ExcelPackage package, ExcelWorksheet worksheet)
        {
            var rowCount = worksheet.Dimension?.Rows;
            var colCount = worksheet.Dimension?.Columns;

            if (!rowCount.HasValue || !colCount.HasValue)
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            for (int row = 1; row <= rowCount.Value; row++)
            {
                for (int col = 1; col <= colCount.Value; col++)
                {
                    sb.AppendFormat("{0}\t", worksheet.Cells[row, col].Value);
                }
                sb.Append(Environment.NewLine);
            }
            return sb.ToString();
        }

        private ExcelPackage createExcelPackage()
        {
            var package = new ExcelPackage();
            package.Workbook.Properties.Title = "Salary Report";
            package.Workbook.Properties.Author = "Vahid N.";
            package.Workbook.Properties.Subject = "Salary Report";
            package.Workbook.Properties.Keywords = "Salary";


            var worksheet = package.Workbook.Worksheets.Add("Employee");

            //First add the headers
            worksheet.Cells[1, 1].Value = "Product Code";
            worksheet.Cells[1, 2].Value = "SKU";
            worksheet.Cells[1, 3].Value = "Title";
            worksheet.Cells[1, 4].Value = "Short_Description";
            worksheet.Cells[1, 5].Value = "Barcode";
            worksheet.Cells[1, 6].Value = "Is Variation Group";
            worksheet.Cells[1, 7].Value = "Variation Group Name";
            worksheet.Cells[1, 8].Value = "Variation SKU";
            worksheet.Cells[1, 9].Value = "Category";
            worksheet.Cells[1, 10].Value = "Purchase Price";
            worksheet.Cells[1, 11].Value = "Retail Price";
            worksheet.Cells[1, 12].Value = "Tax Rate";
            worksheet.Cells[1, 13].Value = "Weight";
            worksheet.Cells[1, 14].Value = "Location";
            worksheet.Cells[1, 15].Value = "Stock Level";
            worksheet.Cells[1, 16].Value = "Primary Image";
            worksheet.Cells[1, 17].Value = "Image 1";
            worksheet.Cells[1, 18].Value = "Image 2";
            worksheet.Cells[1, 19].Value = "Image 3";
            worksheet.Cells[1, 20].Value = "Image 4";
            worksheet.Cells[1, 21].Value = "Brand";
            worksheet.Cells[1, 22].Value = "SubCategory";
            worksheet.Cells[1, 23].Value = "Size";
            worksheet.Cells[1, 24].Value = "Size Map";
            worksheet.Cells[1, 25].Value = "Colour";
            worksheet.Cells[1, 26].Value = "Colour Listed";
            worksheet.Cells[1, 27].Value = "Gender";
            worksheet.Cells[1, 28].Value = "Fabric Compostion";
            worksheet.Cells[1, 29].Value = "Origin";
            worksheet.Cells[1, 30].Value = "RRP";
            worksheet.Cells[1, 31].Value = "Search Term 1";
            worksheet.Cells[1, 32].Value = "Search Term 2";
            worksheet.Cells[1, 33].Value = "Search Term 3";
            worksheet.Cells[1, 34].Value = "Search Term 4";
            worksheet.Cells[1, 35].Value = "Search Term 5";
            worksheet.Cells[1, 36].Value = "Key Product Feature 1";
            worksheet.Cells[1, 37].Value = "Key Product Feature 2";
            worksheet.Cells[1, 38].Value = "Key Product Feature 3";
            worksheet.Cells[1, 39].Value = "Key Product Feature 4";
            worksheet.Cells[1, 40].Value = "Key Product Feature 5";
            worksheet.Cells[1, 41].Value = "Long Description";
            worksheet.Cells[1, 44].Value = "";


            //Add values

            var numberformat = "#,##0";
            var dataCellStyleName = "TableNumber";
            var numStyle = package.Workbook.Styles.CreateNamedStyle(dataCellStyleName);
            numStyle.Style.Numberformat.Format = numberformat;

            worksheet.Cells[2, 1].Value = 1000;
            worksheet.Cells[2, 2].Value = "Jon";
            worksheet.Cells[2, 3].Value = "M";
            worksheet.Cells[2, 4].Value = 5000;
            worksheet.Cells[2, 4].Style.Numberformat.Format = numberformat;

            // AutoFitColumns
            worksheet.Cells[1, 1, 4, 4].AutoFitColumns();

            worksheet.HeaderFooter.OddFooter.InsertPicture(
                new FileInfo(Path.Combine(_hostingEnvironment.WebRootPath, "images", "captcha.jpg")),
                PictureAlignment.Right);

            return package;
        }
    }
}