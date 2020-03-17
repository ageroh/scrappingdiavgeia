using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SoberScrappingTool.Models;
using SoberScrappingTool.Services;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace SoberScrappingTool
{
    class Program
    {
        private static string DateFormatWanted = "dd/MM/yyyy HH:mm:ss";
        private static string SearchKeywordString = ConfigurationManager.AppSettings["searchKeywords"];


        static void Main(string[] args)
        {
            var allKeywords = SetupKeywords(args);
            var emailService = new EmailService(ConfigurationManager.AppSettings);

            var results = new List<ScrappedResults>();
            var scrappingService = new DiageiaScrappingService();

            try
            {
                var excelFile = OpenOrCreateDefaultExcel();
                if (!File.Exists(excelFile.FullName))
                {
                    InitExcelFile(excelFile);
                }

                foreach (var searchKeyword in allKeywords)
                {
                    Console.WriteLine($"Start searching for {searchKeyword}");

                    var customExcelRows = scrappingService.SearchForKeyword(searchKeyword);

                    var scrappedResult = SearchAndScrapResults(customExcelRows, excelFile, searchKeyword);

                    results.Add(scrappedResult);
                }

                AddOrUpdateExcelResults(results, excelFile);

                //count and send email.
                if (results.Any(z => z.DataRowsToUpdate.Any()) || results.Any(z => z.DataRowsToAdd.Any()))
                {
                    // fix the format of body.
                    emailService.SendEmail(
                        $"We got new incoming for some keywords: <br/>" +
                        $"New found={results.Sum(z => z.DataRowsToAdd.Count)} <br/>" +
                        $"Updated total={results.Sum(z => z.DataRowsToUpdate.Count)} <br/>" +
                        $"Searched for keywords: {string.Join(",", allKeywords)}. " +
                        $"Please check the xls document.<br/><br/>Cheers!<br/>");
                }
                else
                {
                    Console.WriteLine($"Nothing new or update found for keywords: {string.Join(",", allKeywords)} at {DateTime.UtcNow.ToShortDateString()}");
                }
            }
            catch (Exception e)
            {
                emailService.SendEmail("Error: " + e.ToString(), "Some error occurred while scrapping.");
            }

            Console.WriteLine("");
            Console.WriteLine($"Copyleft {DateTime.UtcNow.Year}: outofmemo");
        }



        private static void AddOrUpdateExcelResults(List<ScrappedResults> results, FileInfo excelInfo)
        {
            foreach(var result in results)
            {
                if(result.DataRowsToAdd?.Any() ?? false)
                {
                    AddResultToExcel(excelInfo, result.DataRowsToAdd, result.SearchKeyword);
                }
                if(result.DataRowsToUpdate?.Values?.Any() ?? false)
                {
                    UpdateResultToExcel(excelInfo, result.DataRowsToUpdate, result.SearchKeyword);
                }
            }
        }

        private static List<string> SetupKeywords(string[] args)
        {
            var allKeywords = new List<string>();
            if (args.Length == 1)
            {
                Console.WriteLine($"Will search only for keyword:{args[0]}");
                allKeywords.Add(args[0]);
            }
            else
            {
                allKeywords = SearchKeywordString?.Split(';').ToList();
                if (allKeywords != null && allKeywords.Any())
                {
                    Console.WriteLine("Start searching for keywords found in configuration...");
                }

            }

            return allKeywords;
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

                var lastRow = GetExcelLastEmptyRowNumber(worksheet);

                worksheet.Cells[lastRow, 1].LoadFromArrays(allCellData);
                excel.Save();
            }
        }

        private static int GetExcelLastEmptyRowNumber(ExcelWorksheet worksheet)
        {
            int i = 1;
            while (i < 10000)
            {
                if (string.IsNullOrWhiteSpace(worksheet.Cells[i, 1]?.Value?.ToString()))
                {
                    return i;
                }
                i++;
            }
            return i;
        }


        private static ScrappedResults SearchAndScrapResults(IEnumerable<CustomExcelRow> customExcelRows, FileInfo excelFile, string keyword)
        {

            var dataRowsToUpdate = new Dictionary<int, CustomExcelRow>();
            var dataRowsToAdd = new List<CustomExcelRow>();

            // todo: make this happen in SearchInExcel not as a foreach to avoid i/o.
            foreach (var excelRowObject in customExcelRows)
            {
                (int row, Status status) = SearchInExcel(excelRowObject.Code, excelFile, excelRowObject.LastDateChanged);
                if (status == Status.Add)
                {
                    dataRowsToAdd.Add(excelRowObject);
                }
                if (status == Status.Update)
                {
                    dataRowsToUpdate.Add(row, excelRowObject);
                }
            }

            return new ScrappedResults
            {
                DataRowsToAdd = dataRowsToAdd,
                DataRowsToUpdate = dataRowsToUpdate,
                SearchKeyword = keyword,
            };
        }

        private static (int, Status) SearchInExcel(string code, FileInfo excelFile, DateTime lastDateChanged)
        {
            using (ExcelPackage excel = new ExcelPackage(excelFile))
            {
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                var foundRow = worksheet.Cells["A:A"]
                                .Where(cell => cell.Text.ToString().Replace(" ", "") == code.Replace(" ", ""))
                                .Select(z => z.Start.Row)
                                .FirstOrDefault();

                if (foundRow > 0)
                {
                    if (!DateTime.TryParseExact(worksheet.Cells[foundRow, 4].Text, DateFormatWanted, null, System.Globalization.DateTimeStyles.None, out var date))
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
}
