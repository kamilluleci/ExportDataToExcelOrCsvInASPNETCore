using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        List<Author> authors = new List<Author>
        {
            new Author { Id = 1, FirstName = "Joydip", LastName = "Kanjilal" },
            new Author { Id = 2, FirstName = "Steve", LastName = "Smith" },
            new Author { Id = 3, FirstName = "Anand", LastName = "Narayaswamy"}
        };

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult DownloadCommaSeperatedFile()
        {
            try
            {
                StringBuilder stringBuilder = new StringBuilder();
                stringBuilder.AppendLine("Id,FirstName,LastName");
                foreach (var author in authors)
                {
                    stringBuilder.AppendLine($"{author.Id},{ author.FirstName},{ author.LastName}");
                }
                return File(Encoding.UTF8.GetBytes
                    (stringBuilder.ToString()), "text/csv", "authors.csv");
            }
            catch
            {
                return Error();
            }
        }

        public IActionResult DownloadExcelDocument()
        {
            string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            string fileName = "authors.xlsx";
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    IXLWorksheet worksheet =
                        workbook.Worksheets.Add("Authors");
                    worksheet.Cell(1, 1).Value = "Id";
                    worksheet.Cell(1, 2).Value = "FirstName";
                    worksheet.Cell(1, 3).Value = "LastName";
                    for (int index = 1; index <= authors.Count; index++)
                    {
                        worksheet.Cell(index + 1, 1).Value =
                            authors[index - 1].Id;
                        worksheet.Cell(index + 1, 2).Value =
                            authors[index - 1].FirstName;
                        worksheet.Cell(index + 1, 3).Value =
                            authors[index - 1].LastName;
                    }
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();
                        return File(content, contentType, fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                return Error();
            }
        }
    }
}
