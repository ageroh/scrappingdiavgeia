using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;

namespace SoberScrappingTool
{
    class Program
    {
        private static string DateFormatWanted = "dd/MM/yyyy HH:mm:ss";
        
        static void Main(string[] args)
        {
            // using this...
            IWebDriver webDriver = new ChromeDriver();
            
            if(args.Length != 1)
            {
                Console.WriteLine("Enter only one keyword to search for...");
                return;
            }
            var searchKeyword = args[0];


            webDriver.Url = $@"https://diavgeia.gov.gr/search?query=q:%22{searchKeyword}%22&page=0&sort=recent";

            var ispageLoad = new WebDriverWait(webDriver, TimeSpan.FromSeconds(120))
                .Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));

            if (ispageLoad)
                Console.WriteLine("Page completely loaded");
            else
            {
                Console.WriteLine("Page could not be loaded in 2minutes, maybe site is down.");
                return;
            }
            Thread.Sleep(10000);
            var results = webDriver.FindElements(By.CssSelector(".row-fluid.search-rows"));
            if(results?.Count <= 0)
            {
                Console.WriteLine($"Nothing was found at {DateTime.Now} for keyword:{searchKeyword}");
                return;
            }
            Console.WriteLine($"Found total of {results?.Count} in Page 1 of diavgeia for keyword:{searchKeyword}");


            var excelFile = OpenOrCreateDefaultExcel();

            if (!File.Exists(excelFile.FullName))
            {
                InitExcelFile(excelFile);
            }

            var dataRowsToUpdate = new Dictionary<int, CustomExcelRow>();
            var dataRowsToAdd = new List<CustomExcelRow>();
            foreach (var result in results)
            {
                var excelRowObject = ExtraExcelRowForResult(result);
                if(excelRowObject == null)
                {
                    continue;
                }

                (int row, Status status) = SearchInExcel(excelRowObject.Code, excelFile, excelRowObject.LastDateChanged);
                if(status == Status.Add)
                {
                    dataRowsToAdd.Add(excelRowObject);
                }
                if(status == Status.Update)
                {
                    dataRowsToUpdate.Add(row, excelRowObject);
                }
            }

            AddResultToExcel(excelFile, dataRowsToAdd, searchKeyword);

            // peding is the update.
            UpdateResultToExcel(excelFile, dataRowsToUpdate, searchKeyword);

            Console.WriteLine($"New Praxis found:{dataRowsToAdd.Count}");
            Console.WriteLine($"Updated a Praxis with changed date:{dataRowsToUpdate.Count}");
            Console.WriteLine("");
            Console.WriteLine("Copyright: outofmemo");
        }

        private static void UpdateResultToExcel(FileInfo excelFile, Dictionary<int, CustomExcelRow> dataRowsToUpdate, string searchKeyword)
        {
            using (ExcelPackage excel = new ExcelPackage(excelFile))
            {
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];


                List<object[]> allCellData = new List<object[]>();
                foreach (var row in dataRowsToUpdate)
                {
                    var data = row.Value;
                    worksheet.Cells[row.Key, 1].Value = data.Code;
                    worksheet.Cells[row.Key, 2].Value = data.Provider;
                    worksheet.Cells[row.Key, 3].Value = data.Description;
                    worksheet.Cells[row.Key, 4].Value = data.LastDateChanged.ToString(DateFormatWanted);
                    worksheet.Cells[row.Key, 5].Value = data.Type;
                    worksheet.Cells[row.Key, 6].Value = data.Categories;
                    worksheet.Cells[row.Key, 7].Value = data.Pdf;
                    worksheet.Cells[row.Key, 8].Value = searchKeyword;
                    //"Code", "Provider", "Description", "Last Date Changed", "Type", "Categories", "Pdf", keyword
                }
                excel.Save();
            }
        }

        private static void AddResultToExcel(FileInfo excelInfo, List<CustomExcelRow> dataRows, string searchKeyword)
        {
            using (ExcelPackage excel = new ExcelPackage(excelInfo))
            {
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                List<object[]> allCellData = new List<object[]>();
                foreach (var cellData in dataRows)
                {
                    allCellData.Add(new object[] {
                        cellData.Code
                        , cellData.Provider
                        , cellData.Description
                        , cellData.LastDateChanged.ToString(DateFormatWanted)
                        , cellData.Type
                        , cellData.Categories 
                        , cellData.Pdf
                        , searchKeyword
                        //"Code", "Provider", "Description", "Last Date Changed", "Type", "Categories", "Pdf"
                    });
                }

                worksheet.Cells[2, 1].LoadFromArrays(allCellData);
                excel.Save();
            }
        }

        private static CustomExcelRow ExtraExcelRowForResult(IWebElement result)
        {
            // search if already exists in file.
            if (!result.Text.Contains("Ημ/νία τελευταίας τροποποίησης:"))
            {
                Console.WriteLine("Not found excpected date format, missing Ημ/νία τελευταίας τροποποίησης...");
                return null;
            }

            if(!DateTime.TryParseExact(
                result.Text.Substring(result.Text.IndexOf("Ημ/νία τελευταίας τροποποίησης:") + 32, 19)
                , DateFormatWanted
                ,null
                , System.Globalization.DateTimeStyles.None, out var lastDateChanged))
            {
                Console.WriteLine("Not correct date. for result: " + result.Text);
                return null;
            }

            string pattern = @"[Α-Ζ]+\: [0-9α-ωΑ-Ωa-fA-F]{4,14}-[0-9α-ωΑ-Ωa-fA-F]{1,4}";
            var patternMatchCode = Regex.Match(result.Text, pattern);
            if (!patternMatchCode.Success)
            {
                Console.WriteLine($"For result {result.TagName} we cannot find code for document.");
                return null;
            }

            var descriptionWithCode = result.FindElement(By.CssSelector("a[title='Προβολή πράξης']")).Text;
            var descriptionWithoutCode = descriptionWithCode.Substring(descriptionWithCode.IndexOf(" - ") + 3);

            var providerLink = result.FindElement(By.CssSelector("a[title*='Μετάβαση στη σελίδα του φορέα']"));
            var providerLinkName = providerLink.Text;

            var type = result.FindElement(By.CssSelector("a[title*='Αναζήτηση στο είδος πράξης']"));
            var TypeText = type.Text;

            var pdfLink = result.FindElement(By.CssSelector("a[title='Λήψη αρχείου']"));
            var pdf = pdfLink.GetAttribute("href");


            // "Code", "Provider", "Description", "Last Date Changed", "Type", "Categories", "Pdf"

            return new CustomExcelRow
            {
                Code = patternMatchCode.Value,
                LastDateChanged = lastDateChanged,
                Description = descriptionWithoutCode,
                Categories = string.Empty,
                Pdf = pdf,
                Provider = providerLinkName,
                Type = TypeText
            };

        }

        private static (int, Status) SearchInExcel(string code, FileInfo excelFile, DateTime lastDateChanged)
        {
            using (ExcelPackage excel = new ExcelPackage(excelFile))
            {
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                var foundRow = worksheet.Cells["A:A"]
                                .Where(cell => cell.Value.ToString() == code)
                                .Select(z => z.Start.Row)
                                .FirstOrDefault();

                if (foundRow > 0)
                {
                    if(!DateTime.TryParseExact(worksheet.Cells[foundRow, 4].Text, DateFormatWanted, null, System.Globalization.DateTimeStyles.None, out var date ))
                    {
                        return (-1, Status.DontAddNorUpdate);
                    }
                    if (date != lastDateChanged)
                    {
                        return (foundRow, Status.Update);
                    }
                    return (foundRow, Status.DontAddNorUpdate);
                }
            }
            return (0, Status.Add);
        }

        private static FileInfo OpenOrCreateDefaultExcel()
        {
            FileInfo excelFile;
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");

                excelFile = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"..\\diavgeia_{DateTime.UtcNow.Year}.xlsx"));

                if (!excelFile.Exists)
                {
                    excel.SaveAs(excelFile);
                }
            }
            return excelFile;
        }

        private static void InitExcelFile(FileInfo excelInfo)
        {
            using (ExcelPackage excel = new ExcelPackage(excelInfo))
            {
                var headerRow = new List<string[]>()
                  {
                    new string[] { "Code", "Provider", "Description", "Last Date Changed", "Type", "Categories", "Pdf", "SearchKeyword" }
                  };

                // Determine the header range (e.g. A1:D1)
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                // Target a worksheet
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                // Popular header row data
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                excel.SaveAs(excelInfo);
            }

        }
    }

    internal enum Status
    {
        DontAddNorUpdate,
        DontAdd,
        Update,
        Add
    }

    public class CustomExcelRow
    {
        public string Code { get; set; }
        public DateTime LastDateChanged { get; set; }
        public string Pdf { get; set; }
        public string Provider { get; set; }
        public string Description { get; set; }
        public string Type { get; set; }
        public string Categories { get; set; }
    }
}
