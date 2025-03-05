using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Runtime.InteropServices.ComTypes;
using HtmlAgilityPack;
using System.Net;
using Microsoft.Identity.Client;
using System.Security;
using MimeKit.IO;
using System.Text.RegularExpressions;
using Azure.Identity;
using System.Runtime.CompilerServices;
using Azure.Core;
using FMSoftlab.Collaboration.Mail;
using Microsoft.Extensions.Logging;
using FMSoftlab.Logging;

namespace FMSoftlab.Collaboration.Mail.MailKit
{

    public class BaseEmailClient
    {
        public IMailServerConnectionSettings ConnectionSettings { get; }
        protected readonly ILogger _log;
        public BaseEmailClient(IMailServerConnectionSettings connectionSettings, ILogger log)
        {
            _log = log;
            ConnectionSettings = connectionSettings;
            log?.LogDebug($"Server:{connectionSettings?.Server}");
            log?.LogDebug($"Port:{connectionSettings?.Port}");
            log?.LogDebug($"Domain:{connectionSettings?.Domain}");
            log?.LogDebug($"UserName:{connectionSettings?.UserName}");
            log?.LogDebug($"Password:{connectionSettings?.Password}");
            log?.LogDebug($"SmtpAuthenticationFlow:{connectionSettings?.SmtpAuthenticationFlow}");
            log?.LogDebug($"SaslMechanismType:{connectionSettings?.SaslMechanismType}");
            log?.LogDebug($"SecureSocketOptionsType:{connectionSettings?.SecureSocketOptionsType}");
        }

        protected NetworkCredential GetCredentials()
        {
            NetworkCredential res = null;
            _log?.LogDebug($"GetCredentials, Getting mail service credentials for username:{ConnectionSettings.UserName}, Domain:{ConnectionSettings.Domain}");
            if (string.IsNullOrWhiteSpace(ConnectionSettings.UserName))
            {
                _log?.LogWarning("GetCredentials, No username, no credentials!");
                return res;
            }
            if (string.IsNullOrWhiteSpace(ConnectionSettings.Domain))
            {
                _log?.LogInformation($"GetCredentials, username:{ConnectionSettings.UserName}, No domain set");
                res = new NetworkCredential(ConnectionSettings.UserName, ConnectionSettings.Password);
                return res;
            }
            _log?.LogInformation($"GetCredentials, username:{ConnectionSettings.UserName}, domain:{ConnectionSettings.Domain}");
            res = new NetworkCredential(ConnectionSettings.UserName, ConnectionSettings.Password, ConnectionSettings.Domain);
            return res;
        }
    }
    public class FmSoftlabImapClient : BaseEmailClient
    {
        public FmSoftlabImapClient(IMailServerConnectionSettings connectionSettings, ILogger log) : base(connectionSettings, log)
        {

        }
        public async Task<ImapClient> GetImapClientAsync()
        {
            _log?.LogDebug($"GetImapClientAsync in, {ConnectionSettings.Server}:{ConnectionSettings.Port}, Domain:{ConnectionSettings.Domain}, User:{ConnectionSettings.UserName}");
            ImapClient client = new ImapClient();
            client.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
            await client.ConnectAsync(ConnectionSettings.Server, ConnectionSettings.Port, SecureSocketOptions.Auto);
            ICredentials cred = GetCredentials();
            if (cred != null)
            {
                await client.AuthenticateAsync(cred);
                _log?.LogDebug("client authenticated!");
            }
            else
            {
                _log?.LogWarning("Client connected as anonymous user");
            }
            _log?.LogDebug("GetImapClientAsync out");
            return client;
        }

        public async Task<ImapClient> GetImapClientReadonlyInboxAsync()
        {
            ImapClient client = await GetImapClientAsync();
            client.Inbox.Open(FolderAccess.ReadOnly);
            return client;
        }

        public async Task<MimeMessage> FetchLatest()
        {
            MimeMessage res = null;
            using (var client = await GetImapClientReadonlyInboxAsync())
            {
                var messum = client.Inbox.Fetch(client.Inbox.Count() - 1, -1, //MessageSummaryItems.Full | 
               MessageSummaryItems.UniqueId).ToList().FirstOrDefault();
                res = client.Inbox.GetMessage(messum.UniqueId);
            }
            return res;
        }
    }

    public class FmSoftlabSmtpClient : BaseEmailClient
    {
        public FmSoftlabSmtpClient(IMailServerConnectionSettings connectionSettings, ILogger log) : base(connectionSettings, log)
        {

        }
        private async Task ReportMessageSizeAsync(MimeMessage message)
        {
            //log.Debug("ReportMessageSizeAsync in");
            if (message == null) return;
            try
            {
                using (var stream = new MeasuringStream())
                {
                    await message.WriteToAsync(stream);
                    _log?.LogDebug($"message:{message.MessageId}, size:{stream.Length}");
                }

                if (message.Attachments.Count() > 0)
                {
                    long totalLength = 0;
                    foreach (var attachment in message.Attachments)
                    {
                        using (var stream = new MeasuringStream())
                        {
                            var rfc822 = attachment as MessagePart;
                            var part = attachment as MimePart;
                            if (rfc822 != null)
                            {
                                await rfc822.Message.WriteToAsync(stream);
                            }
                            else
                            {
                                //part.Content.DecodeTo(stream);
                                await part.Content.WriteToAsync(stream);
                            }
                            totalLength += stream.Length;
                            _log?.LogDebug($"message:{message.MessageId}, attachment:{attachment.ContentId}, ContentLocation:{attachment.ContentLocation}, size:{stream.Length}, total attachment size:{totalLength}");
                        }
                    }
                }
                else _log?.LogWarning($"message {message.MessageId} has no attachments, total attachment size:0");
            }
            catch (Exception e)
            {
                _log?.LogAllErrors(e);
            }
            //log.Debug("ReportMessageSizeAsync out");
        }
        private string GenerateMessageId(string domain)
        {
            var timestamp = DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ");
            var uniquePart = Guid.NewGuid().ToString(); // Or a counter, PID, or random value
            return $"<{timestamp}.{uniquePart}@{domain}>";
        }
        public async Task<SendMailResult> SendMessageAsync(MimeMessage message)
        {
            _log?.LogDebug("SendMessageAsync MAILKIT in");
            if (message == null)
            {
                _log?.LogWarning("no message to send!");
                return SendMailResult.Failed("no message to send!");
            }
            try
            {
                await ReportMessageSizeAsync(message);
                using (SmtpClient client = await GetSmtpClientAsync())
                {
                    if (message.From.FirstOrDefault() is MailboxAddress addr)
                    {
                        string blogicaMessageId = GenerateMessageId(addr.Domain);
                        message.Headers.Add("X-BlogicaMessageId", blogicaMessageId);
                    }
                    string mailkitmessageId = message.MessageId;
                    _log?.LogDebug($"will send mailkit MessageId:{mailkitmessageId}");
                    string serverResponse = await client.SendAsync(message);
                    _log?.LogDebug($"mailkit just sent an email, MessageId:{mailkitmessageId}, server responded:{serverResponse}, will disconnect...");
                    client.Disconnect(true);
                    if (serverResponse.Contains("2.6.0"))
                    {
                        string pattern = @"2\.6\.0\s+([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})";
                        Match match = Regex.Match(serverResponse, pattern, RegexOptions.IgnoreCase);
                        if (match.Success && match.Groups.Count > 0)
                        {
                            mailkitmessageId = match.Groups[1].Value;
                            _log?.LogDebug($"SendMessageAsync, MAILKIT, serverid:{mailkitmessageId}");
                        }
                    }
                    return SendMailResult.SuccessWithId(mailkitmessageId, serverResponse);
                }
            }
            catch (Exception e)
            {
                _log?.LogAllErrors(e);
                throw;
            }
        }

        public async Task<uint> GetMaxSizeRestrictionAsync()
        {
            uint res = 0;
            var client = await GetSmtpClientAsync();
            if (client.Capabilities.HasFlag(SmtpCapabilities.Size))
            {
                res = client.MaxSize;
                _log?.LogInformation($"The SMTP server has a size restriction on messages: {res}");
            }
            return res;
        }

        public bool ReportCapabilities(SmtpClient client)
        {
            bool res = false;
            _log?.LogDebug("ReportCapabilities in");
            if (client.Capabilities.HasFlag(SmtpCapabilities.Authentication))
            {
                res = true;
                var mechanisms = string.Join(", ", client.AuthenticationMechanisms);
                _log?.LogInformation($"The SMTP server supports the following SASL mechanisms: {mechanisms}");
            }
            else
            {
                _log?.LogWarning($"The SMTP server does not support authentication");
            }
            if (client.Capabilities.HasFlag(SmtpCapabilities.Size))
                _log?.LogInformation($"The SMTP server has a size restriction on messages: {client.MaxSize}");

            if (client.Capabilities.HasFlag(SmtpCapabilities.Dsn))
                _log?.LogInformation("The SMTP server supports delivery-status notifications.");

            if (client.Capabilities.HasFlag(SmtpCapabilities.EightBitMime))
                _log?.LogInformation("The SMTP server supports Content-Transfer-Encoding: 8bit");

            if (client.Capabilities.HasFlag(SmtpCapabilities.BinaryMime))
                _log?.LogInformation("The SMTP server supports Content-Transfer-Encoding: binary");

            if (client.Capabilities.HasFlag(SmtpCapabilities.UTF8))
                _log?.LogInformation("The SMTP server supports UTF-8 in message headers.");
            _log?.LogDebug("ReportCapabilities out");
            return res;
        }

        /*public async Task AcquireTokenByUsernamePassword()
        {
            //https://www.unifeyed.com/portal/knowledgebase/16/How-do-I-use-Office365-for-SMTP.html
            //https://github.com/jstedfast/MailKit/issues/989
            //https://github.com/AzureAD/microsoft-authentication-library-for-dotnet

            //https://stackoverflow.com/questions/43473858/connect-to-outlook-office-365-imap-using-oauth2/60773366#60773366
            //https://docs.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth
            const string ClientId = "..."; // The "Application (client) ID" from the Azure apps portal
            const string UserName = "..."; // The Office365 account username - user@domain.tld
            SecureString Password = ...;   // The Office365 account password
            const string Mailbox = "..."; // The shared mailbox name - mailbox@domain.tld

            // Get an OAuth token:
            var scopes = new[] { "https://outlook.office365.com/IMAP.AccessAsUser.All" };
            var app = PublicClientApplicationBuilder.Create(ClientId).WithAuthority(AadAuthorityAudience.AzureAdMultipleOrgs).Build();
            var authenticationResult = await app.AcquireTokenByUsernamePassword(scopes, UserName, Password).ExecuteAsync(cancellationToken);

            // Authenticate the IMAP client:
            using var client = new ImapClient();
            await client.ConnectAsync("outlook.office365.com", 993, SecureSocketOptions.Auto);
            await client.AuthenticateAsync(new SaslMechanismOAuth2(Mailbox, authenticationResult.AccessToken));
        }*/

        public async Task<SmtpClient> GetSmtpClientOAuth2FlowAsync()
        {
            SmtpClient client = null;
            if (!string.IsNullOrWhiteSpace(ConnectionSettings.LoggerFilename))
            {
                _log?.LogDebug($"GetAzureSmtpClientAsync, Mailkit, will use protocol logger, file:{ConnectionSettings.LoggerFilename}");
                ProtocolLogger pl = new ProtocolLogger(ConnectionSettings.LoggerFilename);
                client = new SmtpClient(pl);
            }
            else
            {
                _log?.LogDebug($"GetAzureSmtpClientAsync, Mailkit, no protocol logger");
                client = new SmtpClient();
            }
            var credential = new DefaultAzureCredential(); // You can configure the credentials if needed
            var tokenRequestContext = new TokenRequestContext(new[] { "https://outlook.office365.com/.default" });
            var accessToken = await credential.GetTokenAsync(tokenRequestContext);
            await client.ConnectAsync(ConnectionSettings.Server, ConnectionSettings.Port, SecureSocketOptions.StartTls);
            await client.AuthenticateAsync(new SaslMechanismOAuth2(ConnectionSettings.UserName, accessToken.Token));
            return client;
        }

        public async Task<SmtpClient> GetSmtpClientUsernamePasswordFlowAsync()
        {
            SmtpClient client = null;
            if (!string.IsNullOrWhiteSpace(ConnectionSettings.LoggerFilename))
            {
                _log?.LogDebug($"GetNormalSmtpClientAsync, Mailkit, will use protocol logger, file:{ConnectionSettings.LoggerFilename}");
                ProtocolLogger pl = new ProtocolLogger(ConnectionSettings.LoggerFilename);
                client = new SmtpClient(pl);
            }
            else
            {
                _log?.LogDebug($"GetNormalSmtpClientAsync, Mailkit, no protocol logger");
                client = new SmtpClient();
            }
            client.ServerCertificateValidationCallback = (s, c, h, e) => true;
            SecureSocketOptions secoptions;
            switch (ConnectionSettings.SecureSocketOptionsType)
            {
                case SecureSocketOptionsType.SslOnConnect:
                    secoptions = SecureSocketOptions.SslOnConnect;
                    break;
                case SecureSocketOptionsType.StartTls:
                    secoptions = SecureSocketOptions.StartTls;
                    break;
                case SecureSocketOptionsType.StartTlsWhenAvailable:
                    secoptions = SecureSocketOptions.StartTlsWhenAvailable;
                    break;
                case SecureSocketOptionsType.None:
                    secoptions = SecureSocketOptions.None;
                    break;
                case SecureSocketOptionsType.Auto:
                    secoptions = SecureSocketOptions.Auto;
                    break;
                default:
                    secoptions = SecureSocketOptions.Auto;
                    break;
            }
            _log?.LogDebug($"Will connect to {ConnectionSettings.Server}:{ConnectionSettings.Port}, using {secoptions}");
            await client.ConnectAsync(ConnectionSettings.Server, ConnectionSettings.Port, secoptions);
            bool shouldAuthenticate = ReportCapabilities(client);
            NetworkCredential cred = GetCredentials();
            if (cred != null && shouldAuthenticate)
            {
                if (ConnectionSettings.SaslMechanismType != SaslMechanismType.None)
                {
                    SaslMechanism mech = null;
                    switch (ConnectionSettings.SaslMechanismType)
                    {
                        case SaslMechanismType.NTLM:
                            client.AuthenticationMechanisms.Clear();
                            client.AuthenticationMechanisms.Add("NTLM");
                            mech = new SaslMechanismNtlm(cred);
                            break;
                        case SaslMechanismType.Plain:
                            client.AuthenticationMechanisms.Clear();
                            client.AuthenticationMechanisms.Add("PLAIN");
                            mech = new SaslMechanismPlain(cred);
                            break;
                        case SaslMechanismType.Login:
                            client.AuthenticationMechanisms.Clear();
                            client.AuthenticationMechanisms.Add("LOGIN");
                            mech = new SaslMechanismLogin(cred);
                            break;
                        case SaslMechanismType.OAuth2:
                            client.AuthenticationMechanisms.Clear();
                            client.AuthenticationMechanisms.Add("XOAUTH2");
                            mech = new SaslMechanismOAuth2(cred);
                            break;
                    }
                    _log?.LogInformation($"Will authenticate using SaslMechanism:{mech.MechanismName}");
                    await client.AuthenticateAsync(mech);
                }
                else
                {
                    _log?.LogInformation($"Will authenticate using credentials");
                    await client.AuthenticateAsync(cred);
                }
                _log?.LogInformation("client authenticated!");
            }
            else
            {
                _log?.LogWarning("Client connected as anonymous user");
            }
            _log?.LogDebug("GetNormalSmtpClientAsync out");
            return client;
        }
        public async Task<SmtpClient> GetSmtpClientAsync()
        {
            if (ConnectionSettings.SmtpAuthenticationFlow==SmtpAuthenticationFlow.OAuth2)
            {
                return await GetSmtpClientOAuth2FlowAsync();
            }
            else
            {
                return await GetSmtpClientUsernamePasswordFlowAsync();
            }
        }
    }

    public interface ISmtpMailkitEMailSender
    {
        Task<SendMailResult> SendHtmlMailAsync(
                    string from,
                    string fromAddress,
                    string sender,
                    string senderAddress,
                    string replyTo,
                    string replyToAddress,
                    string to,
                    string toAddress,
                    string cc,
                    string ccAddress,
                    string subject,
                    string htmlBody
            , IEnumerable<FileAttachment> attachments);

        Task<SendMailResult> SendHtmlMailAsync(
                    string from,
                    string fromAddress,
                    string sender,
                    string senderAddress,
                    string replyTo,
                    string replyToAddress,
                    IEnumerable<MailboxAddress> toRecipientList,
                    IEnumerable<MailboxAddress> ccRecipientList,
                    string subject,
                    string htmlBody
            , IEnumerable<FileAttachment> attachments);

        Task<SendMailResult> SendHtmlMailAsync(ISendMailAttributes sendMailAttributes);
    }
    public class SmtpMailkitEMailSender : ISmtpMailkitEMailSender
    {
        private FmSoftlabSmtpClient SmtpClient { get; }
        protected readonly ILogger _log;
        public SmtpMailkitEMailSender(FmSoftlabSmtpClient smtpClient, ILogger log)
        {
            _log = log;
            SmtpClient = smtpClient;
        }
        private MimeMessage GetMailMessage(
                string from,
                string fromAddress,
                string sender,
                string senderAddress,
                string replyTo,
                string replyToAddress,
                string to,
                string toAddress,
                string cc,
                string ccAddress,
                string subject,
                string htmlBody,
                IEnumerable<FileAttachment> attachments
            )
        {
            MimeMessage sendmes = new MimeMessage();
            sendmes.Subject = subject;
            MailboxAddress mbfrom = null;
            MailboxAddress mbto = null;
            MailboxAddress mbReplyTo = null;
            MailboxAddress mbCc = null;
            MailboxAddress mbSender = null;

            Func<string, string, MailboxAddress> getMailAddress = (address, name) =>
            {
                MailboxAddress res = null;
                if (!string.IsNullOrWhiteSpace(address))
                {
                    res = new MailboxAddress(name, address);
                }
                return res;
            };
            mbfrom = getMailAddress(fromAddress, from);
            mbto = getMailAddress(toAddress, to);
            mbReplyTo = getMailAddress(replyToAddress, replyTo);
            mbCc = getMailAddress(ccAddress, cc);
            mbSender = getMailAddress(senderAddress, sender);
            if (mbfrom != null)
            {
                sendmes.From.Add(mbfrom);
            }
            if (mbto != null)
            {
                sendmes.To.Add(mbto);
            }
            if (mbReplyTo != null)
            {
                sendmes.ReplyTo.Add(mbReplyTo);
            }
            if (mbCc != null)
            {
                sendmes.Cc.Add(mbCc);
            }
            if (mbSender != null)
            {
                sendmes.Sender = mbSender;
            }
            /*var text = new TextPart("html");
            text.ContentId = MimeUtils.GenerateMessageId();
            text.Text = htmlBody;
            var related = new MultipartRelated
            {
                Root = text
            };
            Multipart mulitpart = new Multipart("mixed");
            mulitpart.Add(related);
            if (attachments != null)
            {
                foreach (var attachment in attachments)
                {
                    var at = MailKitAttachmentFactory.GetAttachment(attachment);
                    mulitpart.Add(at);
                }
            }*/
            BodyBuilder bb = new BodyBuilder();
            HtmlToText htt = new HtmlToText();
            string plaintext = htt.ConvertHtml(htmlBody);
            bb.TextBody = plaintext;
            bb.HtmlBody = htmlBody;
            //log.Debug(plaintext);
            //log.Debug(htmlBody);
            if (attachments?.Count() > 0)
            {
                _log?.LogDebug($"attachment count:{attachments?.Count()}");
                foreach (var attachment in attachments)
                {
                    string contentType = MimeTypes.GetMimeType(attachment.OriginalFilename);
                    if (!string.IsNullOrWhiteSpace(contentType))
                    {
                        var mimetypes = contentType.Split('/');
                        MimeKit.ContentType ctype = new MimeKit.ContentType(mimetypes[0], mimetypes[1]);
                        bb.Attachments.Add(attachment.OriginalFilename, attachment.Content, ctype);
                        _log?.LogDebug($"added attachment:{attachment.OriginalFilename}");
                    }
                }
            }
            else
            {
                _log?.LogDebug($"no attachments");
            }
            sendmes.Body = bb.ToMessageBody();
            return sendmes;
        }

        private MimeMessage GetMailMessage(
            string from,
            string fromAddress,
            string sender,
            string senderAddress,
            string replyTo,
            string replyToAddress,
            IEnumerable<MailboxAddress> toRecipientList,
            IEnumerable<MailboxAddress> ccRecipientList,
            string subject,
            string htmlBody,
            IEnumerable<FileAttachment> attachments)
        {
            MimeMessage sendmes = new MimeMessage();
            sendmes.Subject = subject;
            MailboxAddress mbfrom = null;
            MailboxAddress mbReplyTo = null;
            MailboxAddress mbSender = null;

            Func<string, string, MailboxAddress> getMailAddress = (address, name) =>
            {
                MailboxAddress res = null;
                if (!string.IsNullOrWhiteSpace(address))
                {
                    res = new MailboxAddress(name, address);
                }
                return res;
            };
            mbfrom = getMailAddress(fromAddress, from);
            mbReplyTo = getMailAddress(replyToAddress, replyTo);
            mbSender = getMailAddress(senderAddress, sender);
            if (mbfrom != null)
            {
                sendmes.From.Add(mbfrom);
            }
            if (mbSender != null)
            {
                sendmes.Sender = mbSender;
            }
            if (mbReplyTo != null)
            {
                sendmes.ReplyTo.Add(mbReplyTo);
            }
            if (toRecipientList?.Count() > 0)
            {
                sendmes.To.AddRange(toRecipientList);
            }
            if (ccRecipientList?.Count() > 0)
            {
                sendmes.Cc.AddRange(ccRecipientList);
            }

            /*var text = new TextPart("html");
            text.ContentId = MimeUtils.GenerateMessageId();
            text.Text = htmlBody;
            var related = new MultipartRelated
            {
                Root = text
            };
            Multipart mulitpart = new Multipart("mixed");
            mulitpart.Add(related);
            if (attachments != null)
            {
                foreach (var attachment in attachments)
                {
                    var at=MailKitAttachmentFactory.GetAttachment(attachment);
                    mulitpart.Add(at);
                }
            }
            sendmes.Body = mulitpart;
            return sendmes;*/

            BodyBuilder bb = new BodyBuilder();
            HtmlToText htt = new HtmlToText();
            string plaintext = htt.ConvertHtml(htmlBody);
            bb.TextBody = plaintext;
            bb.HtmlBody = htmlBody;
            //log.Debug(plaintext);
            //log.Debug(htmlBody);
            if (attachments?.Count() > 0)
            {
                _log?.LogDebug($"attachment count:{attachments?.Count()}");
                foreach (var attachment in attachments)
                {
                    string contentType = MimeTypes.GetMimeType(attachment.OriginalFilename);
                    if (!string.IsNullOrWhiteSpace(contentType))
                    {
                        var mimetypes = contentType.Split('/');
                        MimeKit.ContentType ctype = new MimeKit.ContentType(mimetypes[0], mimetypes[1]);
                        bb.Attachments.Add(attachment.OriginalFilename, attachment.Content, ctype);
                        _log?.LogDebug($"added attachment:{attachment.OriginalFilename}");
                    }
                }
            }
            else
            {
                _log?.LogDebug($"no attachments");
            }
            sendmes.Body = bb.ToMessageBody();
            return sendmes;
        }

        public async Task<SendMailResult> SendHtmlMailAsync(
            string from,
            string fromAddress,
            string sender,
            string senderAddress,
            string replyTo,
            string replyToAddress,
            string to,
            string toAddress,
            string cc,
            string ccAddress,
            string subject,
            string htmlBody
            , IEnumerable<FileAttachment> attachments)
        {
            MimeMessage sendmes = GetMailMessage(
                from,
                fromAddress,
                sender,
                senderAddress,
                replyTo,
                replyToAddress,
                to,
                toAddress,
                cc,
                ccAddress,
                subject,
                htmlBody
                , attachments);
            return await SmtpClient.SendMessageAsync(sendmes);
        }

        public async Task<SendMailResult> SendHtmlMailAsync(
            string from,
            string fromAddress,
            string sender,
            string senderAddress,
            string replyTo,
            string replyToAddress,
            IEnumerable<MailboxAddress> toRecipientList,
            IEnumerable<MailboxAddress> ccRecipientList,
            string subject,
            string htmlBody
            , IEnumerable<FileAttachment> attachments)
        {
            MimeMessage mail = GetMailMessage(
                from,
                fromAddress,
                sender,
                senderAddress,
                replyTo,
                replyToAddress,
                toRecipientList,
                ccRecipientList,
                subject,
                htmlBody, 
                attachments);
            return await SmtpClient.SendMessageAsync(mail);
        }

        public async Task<SendMailResult> SendHtmlMailAsync(string fromAddress, string senderAddress, string replyToAddress, string toAddress, string ccAddress, string subject, string htmlBody
            , IEnumerable<FileAttachment> attachments)
        {
            return await SendHtmlMailAsync(
                string.Empty,
                fromAddress,
                string.Empty,
                senderAddress,
                string.Empty,
                replyToAddress,
                string.Empty,
                toAddress,
                string.Empty,
                ccAddress,
                subject,
                htmlBody
                , attachments);
        }
        public async Task<SendMailResult> SendHtmlMailAsync(ISendMailAttributes sendMailAttributes)
        {
            List<MailboxAddress> toRecipientList = new List<MailboxAddress>();
            List<MailboxAddress> ccRecipientList = new List<MailboxAddress>();
            if (sendMailAttributes?.ToAddresses?.Count() > 0)
            {
                var temptolist =
                    (from ad in sendMailAttributes.ToAddresses
                     select new MailboxAddress(ad.Name, ad.Address)).ToList();
                toRecipientList.AddRange(temptolist);
            }
            if (sendMailAttributes?.CcAddresses?.Count() > 0)
            {
                var tempcclist =
                    (from ad in sendMailAttributes.CcAddresses
                     select new MailboxAddress(ad.Name, ad.Address)).ToList();
                ccRecipientList.AddRange(tempcclist);
            }
            MimeMessage sendmes = GetMailMessage(
                sendMailAttributes.FromDisplayName,
                sendMailAttributes.FromAddress,
                sendMailAttributes.SenderDisplayName,
                sendMailAttributes.SenderAddress,
                sendMailAttributes.ReplyToDisplayName,
                sendMailAttributes.ReplyToAddress,
                toRecipientList,
                ccRecipientList,
                sendMailAttributes.Subject,
                sendMailAttributes.Body,
                sendMailAttributes.Attachments
                );
            return await SmtpClient.SendMessageAsync(sendmes);
        }
    }
}