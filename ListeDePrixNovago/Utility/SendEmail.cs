using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;

namespace ListeDePrixNovago.PDFTemplate
{
    class SendEmail
    {
        private string server;
        private int port;
        private string username;
        private string password;
        public SendEmail(string server, int port, string username, string password)
        {
            this.server = server;
            this.port = port;
            this.username = username;
            this.password = password;
        }
        public void SendPriceList(string fromAdd, string[] toAdd, string subject, string attachmentFileName, string body = "")
        {
            using (Attachment attachment = new Attachment(attachmentFileName))
            {
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(fromAdd);
                foreach (string add in toAdd)
                {
                    mail.To.Add(new MailAddress(add));
                }
                mail.Subject = subject;
                mail.Attachments.Add(attachment);

                SmtpClient client = new SmtpClient(server, port);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(username, password);
                client.EnableSsl = true;
                client.Send(mail);
            }
        }
    }
}
