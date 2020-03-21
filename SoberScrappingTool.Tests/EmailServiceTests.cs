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
            nvk.Add("emailAddressToSend", "testtest@gmail.com");
            nvk.Add("smtpUsername", "scrapper@test.gr");
            nvk.Add("smtpPassword", "password");
            nvk.Add("smtpHost", "test.gr");
            nvk.Add("smtpPort", "25");

            var emailService = new EmailService(nvk);

            emailService.SendEmail("test body", "test subject");


        }
    }
}
