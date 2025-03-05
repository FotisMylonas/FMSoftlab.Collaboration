using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using SendGrid;
using SendGrid.Helpers.Mail;
using System.Net;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using Microsoft.Extensions.Logging;
using FMSoftlab.Logging;

namespace FMSoftlab.Collaboration.Mail.SendGrid
{
    public interface ISendGridSettings
    {
        string SendGridAPIKey { get; }
    }
    public class SendGridSettings : ISendGridSettings
    {
        public string SendGridAPIKey { get; set; }

        public ISendGridSettings SetSendGridAPIKey(string value)
        {
            return new SendGridSettings(value);
        }
        public SendGridSettings()
        {

        }

        public SendGridSettings(string sendGridAPIKey)
        {
            SendGridAPIKey = sendGridAPIKey;
        }
    }

    public class SendGridMailSender : ISendMail
    {
        private readonly ILogger _log;
        private ISendGridSettings Settings { get; }

        public SendGridMailSender(ILogger log, ISendGridSettings settings)
        {
            Settings = settings;
            _log = log;
        }
        public async Task<SendMailResult> SendMailAsync(ISendMailAttributes sendMailAttributes)
        {
            _log?.LogDebug("SendGridMailSender, SendMailAsync SendGrind In");
            if (sendMailAttributes == null)
                return SendMailResult.Failed("SendGridMailSender, No email info");
            //log.Debug("SendGrind SendMailAsync");
            if (!(sendMailAttributes.ToAddresses?.Any()??false))
            {
                return SendMailResult.Failed("SendGridMailSender, No recipients");
            }
            try
            {
                SendGridMessage mail = new SendGridMessage();
                mail.Subject = sendMailAttributes.Subject;
                foreach (FMEmailAddress toAddress in sendMailAttributes.ToAddresses)
                {
                    EmailAddress to = new EmailAddress(toAddress.Address, toAddress.Name);
                    mail.AddTo(to);
                }
                if (!string.IsNullOrWhiteSpace(sendMailAttributes.Body))
                {
                    mail.HtmlContent = sendMailAttributes.Body;
                    HtmlToText htt = new HtmlToText();
                    mail.PlainTextContent = htt.ConvertHtml(sendMailAttributes.Body);
                }
                if (!string.IsNullOrWhiteSpace(sendMailAttributes.ReplyToAddress))
                {
                    _log?.LogDebug($"replyto address:{sendMailAttributes.ReplyToAddress}, display name:{sendMailAttributes.ReplyToDisplayName}");
                    EmailAddress replytoemail = new EmailAddress(sendMailAttributes.ReplyToAddress, sendMailAttributes.ReplyToDisplayName);
                    mail.SetReplyTo(replytoemail);
                }
                EmailAddress from = new EmailAddress(sendMailAttributes.FromAddress, sendMailAttributes.FromDisplayName);
                mail.From = from;
                _log?.LogDebug($"from address:{mail.From.Email}, display name:{mail.From.Name}");
                if (sendMailAttributes.CcAddresses.Count()>0)
                {
                    foreach (var ccAd in sendMailAttributes.CcAddresses)
                    {
                        _log?.LogDebug("CarbonCopy address:{0}", ccAd.Address);
                        EmailAddress cc = null;
                        if (string.IsNullOrWhiteSpace(ccAd.Name))
                        {
                            cc = new EmailAddress(ccAd.Address);
                        }
                        else
                        {
                            cc = new EmailAddress(ccAd.Address, ccAd.Name);
                        }
                        mail.AddCc(cc);
                    }
                }
                if (sendMailAttributes?.Attachments?.Count() > 0)
                {
                    _log?.LogDebug($"event has {sendMailAttributes?.Attachments?.Count()} file attachment(s)");
                    foreach (FileAttachment mailat in sendMailAttributes?.Attachments)
                    {
                        string fname = Path.GetFileName(mailat.OriginalFilename);
                        _log?.LogDebug($"attaching file {fname}");
                        mail.AddAttachment(fname, Convert.ToBase64String(mailat.Content));
                    }
                }
                SendGridClient client = new SendGridClient(Settings.SendGridAPIKey);
                Response response = await client.SendEmailAsync(mail);
                if (response is null)
                {
                    _log?.LogDebug($"SendEmailAsync:no server response");
                    return SendMailResult.Failed("no server response");
                }
                _log?.LogDebug($"TenantID:{sendMailAttributes.TenantId}, SendGrid StatusCode:{response.StatusCode}");
                if (!(response.StatusCode == HttpStatusCode.OK || response.StatusCode == HttpStatusCode.Accepted || response.StatusCode == HttpStatusCode.Created))
                {
                    return SendMailResult.Failed($"SendGrid responded with status {response.StatusCode}, IsSuccessStatusCode:{response.IsSuccessStatusCode}");
                }
                string messageId = string.Empty;
                if (response.Headers.TryGetValues("X-Message-Id", out var messageIdValues))
                {
                    messageId=messageIdValues.FirstOrDefault();
                }
                string bodyres = response.Body.ReadAsStringAsync().Result;
                _log?.LogError($"TenantID:{sendMailAttributes.TenantId}, SendGrid StatusCode:{response.StatusCode}, X-Message-Id:{messageId}{Environment.NewLine}response.Headers:{response.Headers}{Environment.NewLine}response.Body:{bodyres}");
                return SendMailResult.SuccessWithId(messageId);
            }
            catch (Exception e)
            {
                _log?.LogAllErrors(e);
                throw;
            }
        }
    }
}
