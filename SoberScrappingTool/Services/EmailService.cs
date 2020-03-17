using System;
using System.Collections.Specialized;
using System.Linq;
using System.Net;
using System.Net.Mail;

namespace SoberScrappingTool.Services
{
    public class EmailService 
    {
        private readonly string EmailAddressesToSend;
        private readonly string SmtpUsername;
        private readonly string SmtpPassword;
        private readonly string SmtpHost;
        private readonly int SmtpPort;


        public EmailService(NameValueCollection nameValueCollection)
        {
            EmailAddressesToSend = nameValueCollection["emailAddressToSend"];
            SmtpUsername = nameValueCollection["smtpUsername"];
            SmtpPassword = nameValueCollection["smtpPassword"];
            SmtpHost = nameValueCollection["smtpHost"];
            SmtpPort = Convert.ToInt32(nameValueCollection["smtpPort"]);
        }

        public void SendEmail(string body, string subject = null)
        {
            subject = subject ?? $"Sober Scrapping: {DateTime.Now.Date.ToString("dd/MM/yyyy")}";

            var smtp = new SmtpClient
            {
                Host = SmtpHost,
                Port = SmtpPort,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(SmtpUsername, SmtpPassword),

            };

            using (var message = new MailMessage(SmtpUsername, EmailAddressesToSend)
            {
                Subject = subject,
                Body = body,
                IsBodyHtml = true,
            })
            {
                smtp.Send(message);
            }
        }

    }
}
