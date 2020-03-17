using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SoberScrappingTool.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace SoberScrappingTool.Services
{
    public interface IScrappingService
    {
        IEnumerable<CustomExcelRow> SearchForKeyword(string keyword);
    }

    public class DiageiaScrappingService : IScrappingService
    {
        private static string DateFormatWanted = "dd/MM/yyyy HH:mm:ss";

        public IEnumerable<CustomExcelRow> SearchForKeyword(string keyword)
        {
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("headless");

            using (IWebDriver webDriver = new ChromeDriver(chromeOptions))
            {

                webDriver.Url = $@"https://diavgeia.gov.gr/search?query=q:{keyword}&page=0&sort=recent";

                var ispageLoad = new WebDriverWait(webDriver, TimeSpan.FromSeconds(600))
                    .Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));

                if (!ispageLoad)
                {
                    Console.WriteLine("Page could not be loaded in 240 sec, maybe site is down.");
                    return Enumerable.Empty<CustomExcelRow>();
                }

                Thread.Sleep(20000);
                var results = webDriver.FindElements(By.CssSelector(".row-fluid.search-rows"));

                if (results?.Count <= 0)
                {
                    Console.WriteLine($"Nothing was found at {DateTime.Now} for keyword:{keyword}");
                    return Enumerable.Empty<CustomExcelRow>();
                }
                Console.WriteLine($"Found total of {results?.Count} in Page 1 of diavgeia for keyword:{keyword}");

                return results.Select(webElement => ExtraExcelRowForResult(webElement))
                    .ToList();
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

            if (!DateTime.TryParseExact(
                result.Text.Substring(result.Text.IndexOf("Ημ/νία τελευταίας τροποποίησης:") + 32, 19)
                , DateFormatWanted
                , null
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


    }
}
