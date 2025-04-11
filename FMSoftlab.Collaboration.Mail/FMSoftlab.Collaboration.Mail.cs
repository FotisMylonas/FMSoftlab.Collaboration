using HtmlAgilityPack;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace FMSoftlab.Collaboration.Mail
{
    public enum MessagingRecipientType { From, To, CC, ReplyTo };
    public class SendMailResult
    {
        public bool Success { get; set; }
        public string MessageId { get; set; }
        public string ResultMessage { get; set; }
        public SendMailResult(bool success, string messageId)
        {
            Success = success;
            MessageId = messageId;
            ResultMessage=string.Empty;
        }
        public SendMailResult(bool success, string messageId, string resultMessage)
        {
            Success = success;
            MessageId = messageId;
            ResultMessage=resultMessage;
        }

        public SendMailResult(bool success)
        {
            Success = success;
            MessageId = string.Empty;
            ResultMessage=string.Empty;
        }
        public SendMailResult()
        {
            Success = false;
            MessageId = string.Empty;
            ResultMessage=string.Empty;
        }

        public static SendMailResult SuccessNoId = new SendMailResult(true);
        public static SendMailResult SuccessWithId(string id) => new SendMailResult(true, id);
        public static SendMailResult SuccessWithId(string id, string message) => new SendMailResult(true, id, message);
        public static SendMailResult Failed(string message) => new SendMailResult(false, string.Empty, message);
    }

    public interface ISendMail
    {
        Task<SendMailResult> SendMailAsync(ISendMailAttributes sendMailAttributes);
    }
    public interface IEmailAddress
    {
        string Name { get; set; }
        string Address { get; set; }
    }
    public class FMEmailAddress : IEmailAddress
    {
        public string Name { get; set; }
        public string Address { get; set; }

        public FMEmailAddress(string name, string address)
        {
            Name = name;
            Address = address;
        }
        public FMEmailAddress(string address)
        {
            Address = address;
        }
    }
    public class FileAttachment
    {
        public Guid GlobalId { get; set; }
        public string OriginalFilename { get; set; }
        public byte[] Content { get; set; }
        public string ExtractedText { get; set; }

        public FileAttachment()
        {
            GlobalId = Guid.NewGuid();
        }
        public FileAttachment(string filename) : this()
        {
            OriginalFilename = Path.GetFileName(filename);
            LoadFile(filename);
        }

        public async Task LoadFileAsync(string filename)
        {
            if (File.Exists(filename))
            {
                using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read))
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        await fs.CopyToAsync(memoryStream);
                        Content = memoryStream.ToArray();
                    }
                }
            }
        }
        public void LoadFile(string filename)
        {
            if (File.Exists(filename))
            {
                using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read))
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        fs.CopyTo(memoryStream);
                        Content = memoryStream.ToArray();
                    }
                }
            }
        }
        public MemoryStream GetContentAsStream()
        {
            return new MemoryStream(Content);
        }
    }

    public enum SmtpAuthenticationFlow { UsernamePassword, OAuth2 };
    public enum SaslMechanismType { None, NTLM, Plain, Login, OAuth2 };
    public enum SecureSocketOptionsType { None, Auto, StartTlsWhenAvailable, StartTls, SslOnConnect };

    public interface IMailServerConnectionSettings
    {
        string Domain { get; set; }
        string UserName { get; set; }
        string Password { get; set; }
        int Port { get; set; }
        string Server { get; set; }
        string LoggerFilename { get; set; }
        SmtpAuthenticationFlow SmtpAuthenticationFlow { get; set; }
        SaslMechanismType SaslMechanismType { get; set; }
        SecureSocketOptionsType SecureSocketOptionsType { get; set; }
    }

    public class MailServerConnectionSettings : IMailServerConnectionSettings
    {
        public SmtpAuthenticationFlow SmtpAuthenticationFlow { get; set; }
        public string Domain { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public int Port { get; set; }
        public string Server { get; set; }
        public string LoggerFilename { get; set; }

        public SaslMechanismType SaslMechanismType { get; set; }
        public SecureSocketOptionsType SecureSocketOptionsType { get; set; }
        public MailServerConnectionSettings()
        {

        }
        public MailServerConnectionSettings(
            string server,
            int port,
            string userName,
            string password,
            string domain,
            string loggerFilename,
            SmtpAuthenticationFlow smtpAuthenticationFlow,
            SaslMechanismType saslMechanismType,
            SecureSocketOptionsType secureSocketOptionsType)
        {
            Server = server;
            Port = port;
            UserName = userName;
            Password = password;
            Domain = domain;
            LoggerFilename = loggerFilename;
            SmtpAuthenticationFlow =smtpAuthenticationFlow;
            SaslMechanismType = saslMechanismType;
            SecureSocketOptionsType = secureSocketOptionsType;
        }
        public MailServerConnectionSettings(
            string server,
            int port,
            string userName,
            string password,
            string domain,
            string loggerFilename,
            SaslMechanismType saslMechanismType,
            SecureSocketOptionsType secureSocketOptionsType) : this(
                server,
                port,
                userName,
                password,
                domain,
                loggerFilename,
                SmtpAuthenticationFlow.UsernamePassword,
                saslMechanismType,
                secureSocketOptionsType)
        {

        }
    }

    public interface ISendMailAttributes
    {
        int TenantId { get; }
        string Subject { get; }
        string Body { get; }
        string FromAddress { get; }
        string FromDisplayName { get; }
        string SenderAddress { get; }
        string SenderDisplayName { get; }
        string ReplyToAddress { get; }
        string ReplyToDisplayName { get; }
        List<FMEmailAddress> ToAddresses { get; }
        List<FMEmailAddress> CcAddresses { get; }
        List<FileAttachment> Attachments { get; }
        ISendMailAttributes AddRecipient(string address);
        ISendMailAttributes AddRecipient(string name, string address);
        ISendMailAttributes AddRecipient(IEnumerable<FMEmailAddress> addresses);
        ISendMailAttributes SetSubject(string subject);
        ISendMailAttributes SetBody(string body);
        ISendMailAttributes SetFromAddress(string fromAddress);
        ISendMailAttributes SetFromDisplayName(string fromDisplayName);
        ISendMailAttributes SetReplyToAddress(string replyto);
        ISendMailAttributes SetReplyToDisplayName(string replyToDisplayName);
        ISendMailAttributes SetTenantId(int tenantId);
        ISendMailAttributes AddCarbonCopyEmail(string address);
        ISendMailAttributes AddCarbonCopyEmail(string name, string address);
        ISendMailAttributes AddCarbonCopyEmail(IEnumerable<FMEmailAddress> addresses);
        ISendMailAttributes AddCc(string address);
        ISendMailAttributes AddCc(string name, string address);
        ISendMailAttributes AddCc(IEnumerable<FMEmailAddress> addresses);
        ISendMailAttributes AddAttachment(FileAttachment attachment);
        ISendMailAttributes AddAttachment(string filename);
        ISendMailAttributes AddAttachment(string filename, byte[] content);
        ISendMailAttributes SetCcs(IEnumerable<FMEmailAddress> addresses);
        ISendMailAttributes SetRecipients(IEnumerable<FMEmailAddress> addresses);
        ISendMailAttributes SetCarbonCopyEmails(IEnumerable<FMEmailAddress> addresses);
    }

    public class SendMailAttributes : ISendMailAttributes
    {
        public int TenantId { get; set; }

        public string Subject { get; set; }

        public string Body { get; set; }

        public string FromAddress { get; set; }

        public string FromDisplayName { get; set; }

        public string SenderAddress { get; set; }

        public string SenderDisplayName { get; set; }

        public string ReplyToAddress { get; set; }

        public string ReplyToDisplayName { get; set; }

        public List<FMEmailAddress> ToAddresses { get; }
        public List<FMEmailAddress> CcAddresses { get; }
        public List<FileAttachment> Attachments { get; }
        public SendMailAttributes()
        {
            TenantId = 0;
            ToAddresses = new List<FMEmailAddress>();
            CcAddresses = new List<FMEmailAddress>();
            Attachments = new List<FileAttachment>();
        }

        public SendMailAttributes(
                int tenantId,
                string replyToDisplayName,
                string subject,
                string body,
                string fromAddress,
                string fromDisplayName,
                string senderAddress,
                string senderDisplayName)
        {
            TenantId = tenantId;
            ToAddresses = new List<FMEmailAddress>();
            CcAddresses = new List<FMEmailAddress>();
            Attachments = new List<FileAttachment>();
            ReplyToDisplayName = replyToDisplayName;
            Subject = subject;
            Body = body;
            FromAddress = fromAddress;
            FromDisplayName = fromDisplayName;
            SenderAddress = senderAddress;
            SenderDisplayName = senderDisplayName;
        }
        public ISendMailAttributes AddCarbonCopyEmail(string address)
        {
            FMEmailAddress ad = new FMEmailAddress(address);
            CcAddresses.Add(ad);
            return this;
        }
        public ISendMailAttributes AddCarbonCopyEmail(string name, string address)
        {
            FMEmailAddress ad = new FMEmailAddress(name, address);
            CcAddresses.Add(ad);
            return this;
        }
        public ISendMailAttributes AddCarbonCopyEmail(IEnumerable<FMEmailAddress> addresses)
        {
            if (addresses is null)
                return this;
            if (addresses?.Any() ?? false)
            {
                CcAddresses.AddRange(addresses);
            }
            return this;
        }
        public ISendMailAttributes SetCarbonCopyEmails(IEnumerable<FMEmailAddress> addresses)
        {
            if (addresses?.Any() ?? false)
                AddCarbonCopyEmail(addresses);
            return this;
        }

        public ISendMailAttributes AddRecipient(string address)
        {
            FMEmailAddress ad = new FMEmailAddress(address);
            ToAddresses.Add(ad);
            return this;
        }

        public ISendMailAttributes AddRecipient(string name, string address)
        {
            FMEmailAddress ad = new FMEmailAddress(name, address);
            ToAddresses.Add(ad);
            return this;
        }

        public ISendMailAttributes AddRecipient(IEnumerable<FMEmailAddress> addresses)
        {
            if (addresses is null)
                return this;
            if (addresses?.Any() ?? false)
            {
                ToAddresses.AddRange(addresses);
            }
            return this;
        }
        public ISendMailAttributes SetRecipients(IEnumerable<FMEmailAddress> addresses)
        {
            if (addresses?.Any() ?? false)
                AddRecipient(addresses);
            return this;
        }
        public ISendMailAttributes AddCc(string address)
        {
            FMEmailAddress ad = new FMEmailAddress(address);
            CcAddresses.Add(ad);
            return this;
        }
        public ISendMailAttributes AddCc(string name, string address)
        {
            FMEmailAddress ad = new FMEmailAddress(name, address);
            CcAddresses.Add(ad);
            return this;
        }
        public ISendMailAttributes AddCc(IEnumerable<FMEmailAddress> addresses)
        {
            if (addresses is null)
                return this;
            if (addresses?.Any() ?? false)
            {
                CcAddresses.AddRange(addresses);
            }
            return this;
        }
        public ISendMailAttributes SetCcs(IEnumerable<FMEmailAddress> addresses)
        {
            if (addresses?.Any() ?? false)
                AddCc(addresses);
            return this;
        }
        public ISendMailAttributes SetSubject(string subject)
        {
            Subject = subject;
            return this;
        }
        public ISendMailAttributes SetBody(string body)
        {
            Body = body;
            return this;
        }
        public ISendMailAttributes SetFromAddress(string fromAddress)
        {
            FromAddress = fromAddress;
            return this;
        }

        public ISendMailAttributes SetFromDisplayName(string fromDisplayName)
        {
            FromDisplayName = fromDisplayName;
            return this;
        }

        public ISendMailAttributes SetReplyToAddress(string replyToAddress)
        {
            ReplyToAddress = replyToAddress;
            return this;
        }

        public ISendMailAttributes SetReplyToDisplayName(string replyToDisplayName)
        {
            ReplyToDisplayName = replyToDisplayName;
            return this;
        }

        public ISendMailAttributes SetTenantId(int tenantId)
        {
            TenantId = tenantId;
            return this;
        }

        public ISendMailAttributes AddAttachment(FileAttachment attachment)
        {
            Attachments.Add(attachment);
            return this;
        }
        public ISendMailAttributes AddAttachment(string filename)
        {
            FileAttachment at = new FileAttachment(filename);
            Attachments.Add(at);
            return this;
        }
        public ISendMailAttributes AddAttachment(string filename, byte[] content)
        {
            FileAttachment at = new FileAttachment { Content = content, OriginalFilename = filename };
            Attachments.Add(at);
            return this;
        }
    }

    public class HtmlToText
    {
        public bool IsHTML(string input)
        {
            bool res = false;
            string regexpattern = @"<(br|basefont|hr|input|source|frame|param|area|meta|!--|col|link|option|base|img|wbr|!DOCTYPE).*?>|<(a|abbr|acronym|address|applet|article|aside|audio|b|bdi|bdo|big|blockquote|body|button|canvas|caption|center|cite|code|colgroup|command|datalist|dd|del|details|dfn|dialog|dir|div|dl|dt|em|embed|fieldset|figcaption|figure|font|footer|form|frameset|head|header|hgroup|h1|h2|h3|h4|h5|h6|html|i|iframe|ins|kbd|keygen|label|legend|li|map|mark|menu|meter|nav|noframes|noscript|object|ol|optgroup|output|p|pre|progress|q|rp|rt|ruby|s|samp|script|section|select|small|span|strike|strong|style|sub|summary|sup|table|tbody|td|textarea|tfoot|th|thead|time|title|tr|track|tt|u|ul|var|video).*?<\/\2>";
            Regex rgx = new Regex(regexpattern, RegexOptions.IgnoreCase);
            MatchCollection matches = rgx.Matches(input);
            res = (matches.Count > 0);
            return res;
        }

        public string GetTextIfHtml(string html)
        {
            string res = html;
            if ((!string.IsNullOrWhiteSpace(res)) && (IsHTML(res)))
            {
                res = ConvertHtml(res);
            }
            return res;
        }

        public string ConvertHtml(string html)
        {
            string s = "";
            try
            {
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(html);
                using (StringWriter sw = new StringWriter())
                {
                    ConvertTo(doc.DocumentNode, sw);
                    sw.Flush();
                    s = sw.ToString();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return s;
        }

        private void ConvertTo(HtmlNode node, TextWriter outText)
        {
            string html;
            switch (node.NodeType)
            {
                case HtmlNodeType.Comment:
                    // don't output comments
                    break;

                case HtmlNodeType.Document:
                    ConvertContentTo(node, outText);
                    break;

                case HtmlNodeType.Text:
                    // script and style must not be output
                    string parentName = node.ParentNode.Name;
                    if ((parentName == "script") || (parentName == "style"))
                        break;

                    // get text
                    html = ((HtmlTextNode)node).Text;

                    // is it in fact a special closing node output as text?
                    if (HtmlNode.IsOverlappedClosingElement(html))
                        break;

                    // check the text is meaningful and not a bunch of whitespaces
                    if (html.Trim().Length > 0)
                    {
                        outText.Write(HtmlEntity.DeEntitize(html));
                    }
                    break;

                case HtmlNodeType.Element:
                    switch (node.Name)
                    {
                        case "p":
                            // treat paragraphs as crlf
                            outText.Write("\r\n");
                            break;
                    }

                    if (node.HasChildNodes)
                    {
                        ConvertContentTo(node, outText);
                    }
                    break;
            }
        }

        private void ConvertContentTo(HtmlNode node, TextWriter outText)
        {
            foreach (HtmlNode subnode in node.ChildNodes)
            {
                ConvertTo(subnode, outText);
            }
        }

    }
}
