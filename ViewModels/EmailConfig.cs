namespace MVC_Music.ViewModels
{
    public class EmailConfig : IEmailConfig
    {
        public string SmtpServer { get; set; }
        public int SmtpPort { get; set; }
        public string SmtpFromName { get; set; }
        public string SmtpUsername { get; set; }
        public string SmtpPassword { get; set; }
    }
}
