using System;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Windows;

namespace WPFMailerCore
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static string From { get; set; }
        private static string To { get; set; }
        private static string EmailSubject { get; set; }
        private static string Smtp { get; set; }
        private static int Port { get; set; }
        private static bool UseSsl { get; set; }
        private static bool IsBodyHTml { get; set; }
        private static string HtmlBodyString { get; set; }
        private static string Username { get; set; }
        private static string Password { get; set; }

        public MainWindow()
        {
            InitializeComponent();
        }


        private void SetStatus(string status)
        {
            LbStatus.Text = status;

        }

        private void BtExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void BtClear_Click(object sender, RoutedEventArgs e)
        {
            TbFrom.Text = string.Empty;
            TbTo.Text = string.Empty;
            Tbsmtp.Text = string.Empty;
            TbPassword.Password = string.Empty;
            TbUsername.Text = string.Empty;
            TbSubject.Text = string.Empty;

            TbPort.Text = "587";

            CbSSL.IsChecked = true;
            CbHtmlBody.IsChecked = true;

            TbBody.Text = string.Empty;

        }

        private async void BtSend_Click(object sender, RoutedEventArgs e)
        {
            SetStatus("Working on it");
            int tmp;

            From = TbFrom.Text;
            To = TbTo.Text;
            Smtp = Tbsmtp.Text;
            Username = TbUsername.Text;
            EmailSubject = TbSubject.Text;
            Password = TbPassword.Password;
            HtmlBodyString = TbBody.Text;

            int.TryParse(TbPort.Text, out tmp);
            Port = tmp;

            IsBodyHTml = CbHtmlBody.IsChecked ?? false;
            UseSsl = CbSSL.IsChecked ?? false;

            try
            {
                await Email();
                SetStatus("Email Sent");
            }
            catch (Exception ex)
            {
                SetStatus(
                    $"{ex.Message}\n\n" +
                    "Make Sure that UserName and Password is Correct and your Public IP is not Blacklisted https://mxtoolbox.com/blacklists.aspx \n\n" +
                    "If MultiFactor Authentication is Enabled on your (Gmail or SMTP account) you will need to use an app password instead of your regular password (https://myaccount.google.com/apppasswords)\n" +
                    "Because you don't have the public IP on your Exchange Online Connector\n And the port should be 25 without Username and Password.");

            }
        }

        private static async Task Email()
        {
            MailMessage message = new MailMessage();
            SmtpClient smtp = new SmtpClient();
            message.From = new MailAddress(From);
            message.To.Add(new MailAddress(To));
            message.Subject = EmailSubject;
            message.Body = HtmlBodyString;
            smtp.Port = Port;
            smtp.Host = Smtp;
            smtp.EnableSsl = UseSsl;
            message.IsBodyHtml = IsBodyHTml;

            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.Credentials = (string.IsNullOrEmpty(Username) && string.IsNullOrEmpty(Password))
                ? null
                : new NetworkCredential(userName: Username, Password);
            if (smtp.Credentials == null)
            {
                smtp.UseDefaultCredentials = true;
            }

            await smtp.SendMailAsync(message);

        }


    }
}
