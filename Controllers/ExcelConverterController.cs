using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Task2.Controllers
{
    public class ExcelConverterController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ConvertToSql(IFormFile excelFile, string tableName)
        {
            if (excelFile != null && excelFile.Length > 0 && !string.IsNullOrEmpty(tableName))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(excelFile.OpenReadStream()))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet != null)
                    {
                        var sqlStatements = new List<string>();
                        var rowCount = worksheet.Dimension.Rows;
                        var colCount = worksheet.Dimension.Columns;

                        var headers = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            headers.Add(worksheet.Cells[1, col].Text);
                        }

                        var createTableStatement = $"CREATE TABLE {tableName} (\n";
                        createTableStatement += string.Join(",\n", headers.Select(header => $"{header} NVARCHAR(MAX)"));
                        createTableStatement += "\n);";
                        sqlStatements.Add(createTableStatement);

                        for (int row = 2; row <= rowCount; row++) 
                        {
                            var columns = new List<string>();
                            for (int col = 1; col <= colCount; col++)
                            {
                                var cellValue = worksheet.Cells[row, col].Text;

                                if (IsNumeric(cellValue))
                                {
                                    columns.Add(cellValue); 
                                }
                                else
                                {
                                    var formattedValue = $"N\"{cellValue.Replace("\"", "\"\"")}\"";
                                    columns.Add(formattedValue);
                                }
                            }

                            var insertStatement = $"INSERT INTO {tableName} ({string.Join(", ", headers)}) VALUES ({string.Join(", ", columns)});";
                            sqlStatements.Add(insertStatement);
                        }

                        ViewBag.SqlStatements = sqlStatements;
                    }
                }
            }

            return View("Index");
        }

        private bool IsNumeric(string value)
        {
            return double.TryParse(value, out _);
        }
    }
}
