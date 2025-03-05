using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net;
using FMSoftlab.Collaboration.Mail;
using Microsoft.Extensions.Logging;
using FMSoftlab.Logging;

namespace FMSoftlab.Collaboration.Mail.SMTP
{

    public class SMTPMailSender : ISendMail
    {
        //private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private readonly ILogger _log;
        private readonly SmtpClient client;

        public SMTPMailSender(ILogger log)
        {
            _log = log;
            client = new SmtpClient();
        }
        private string GenerateMessageId(string domain)
        {
            var timestamp = DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ");
            var uniquePart = Guid.NewGuid().ToString(); // Or a counter, PID, or random value
            return $"<{timestamp}.{uniquePart}@{domain}>";
        }

        public async Task<SendMailResult> SendMailAsync(ISendMailAttributes sendMailAttributes)
        {
            _log?.LogDebug("SendMailAsync SMTP in");
            if (!(sendMailAttributes?.ToAddresses?.Any()?? false))
            {
                _log?.LogInformation("SendMailAsync SMTP, No recipients");
                return SendMailResult.Failed("No recipients");
            }

            using (MailMessage mm = new MailMessage())
            {
                mm.IsBodyHtml = true;
                mm.BodyEncoding = UTF8Encoding.UTF8;
                if (!string.IsNullOrWhiteSpace(sendMailAttributes.ReplyToAddress))
                {
                    MailAddress replytoemail = new MailAddress(sendMailAttributes.ReplyToAddress, sendMailAttributes.ReplyToDisplayName, Encoding.BigEndianUnicode);
                    mm.ReplyToList.Add(replytoemail);
                    _log?.LogDebug($"replyto address:{sendMailAttributes.ReplyToAddress}, display name:{sendMailAttributes.ReplyToDisplayName}");
                }
                _log?.LogDebug($"from address:{sendMailAttributes.FromAddress}, display name:{sendMailAttributes.FromDisplayName}");
                MailAddress from = new MailAddress(sendMailAttributes.FromAddress, sendMailAttributes.FromDisplayName, Encoding.BigEndianUnicode);
                mm.From = from;
                foreach (var toaddress in sendMailAttributes.ToAddresses)
                {
                    mm.To.Add(new MailAddress(toaddress.Address, toaddress.Name));
                }

                if (sendMailAttributes.CcAddresses.Count() > 0)
                {
                    foreach (var ccAd in sendMailAttributes.CcAddresses)
                    {
                        _log?.LogDebug(string.Format("CarbonCopy address:{0}", ccAd.Address));
                        MailAddress cc = null;
                        if (string.IsNullOrWhiteSpace(ccAd.Name))
                        {
                            cc = new MailAddress(ccAd.Address);
                        }
                        else
                        {
                            cc = new MailAddress(ccAd.Address, ccAd.Name);
                        }
                        mm.CC.Add(cc);
                    }
                }
                mm.Subject = sendMailAttributes.Subject;
                mm.Body = sendMailAttributes.Body;
                mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                try
                {
                    //string messageId = $"<{Guid.NewGuid()}@example.com>";
                    string messageId = GenerateMessageId(from.Host);
                    mm.Headers.Add("X-BlogicaMessageId", messageId);
                    await client.SendMailAsync(mm);
                    _log?.LogDebug($"email sent!, messageId:{messageId}");
                    return SendMailResult.SuccessWithId(messageId);
                }
                catch (Exception e)
                {
                    _log?.LogAllErrors(e);
                    throw;
                }
            }
        }
        public SendMailResult SendMail(ISendMailAttributes sendMailAttributes)
        {
            return SendMailAsync(sendMailAttributes).Result;
        }
    }
}