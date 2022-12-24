using Microsoft.AspNetCore.Identity.UI.Services;
using MimeKit.Text;
using MimeKit;
using MVC_Music.ViewModels;
using MailKit.Net.Smtp;

namespace MVC_Music.Utilities
{
    public class EmailService
    {
        public class EmailSender : IEmailSender
        {
            private readonly IEmailConfig _emailConfiguration;
            private readonly ILogger<EmailSender> _logger;

            public EmailSender(IEmailConfig emailConfiguration, ILogger<EmailSender> logger)
            {
                _emailConfiguration = emailConfiguration;
                _logger = logger;
            }
            public async Task SendEmailAsync(string email, string subject, string htmlMessage)
            {
                var message = new MimeMessage();
                message.To.Add(new MailboxAddress(email, email));
                message.From.Add(new MailboxAddress(_emailConfiguration.SmtpFromName, _emailConfiguration.SmtpUsername));

                message.Subject = subject;
                message.Body = new TextPart(TextFormat.Html)
                {
                    Text = htmlMessage
                };
                try
                {
                    using var emailClient = new SmtpClient();
                    emailClient.Connect(_emailConfiguration.SmtpServer, _emailConfiguration.SmtpPort, false);
                    emailClient.AuthenticationMechanisms.Remove("XOAUTH2");
                    emailClient.Authenticate(_emailConfiguration.SmtpUsername, _emailConfiguration.SmtpPassword);
                    await emailClient.SendAsync(message);
                    emailClient.Disconnect(true);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex.GetBaseException().Message);
                }
            }
        }
    }
}
