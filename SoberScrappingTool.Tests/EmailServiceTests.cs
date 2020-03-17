using Microsoft.VisualStudio.TestTools.UnitTesting;
using SoberScrappingTool.Services;
using System.Collections.Specialized;

namespace Scrapping.Tests
{
    [TestClass]
    public class EmailServiceTests
    {
        [TestMethod]
        public void TestMethod1()
        {
            var nvk = new NameValueCollection();
            nvk.Add("emailAddressToSend", "argigero@gmail.com,agerogiannis@icloud.com");
            nvk.Add("smtpUsername", "scrapper@megael.gr");
            nvk.Add("smtpPassword", "scr@pp3r");
            nvk.Add("smtpHost", "mail.megael.gr");
            nvk.Add("smtpPort", "25");

            var emailService = new EmailService(nvk);

            emailService.SendEmail("test body", "test subject");

        }
    }
}
