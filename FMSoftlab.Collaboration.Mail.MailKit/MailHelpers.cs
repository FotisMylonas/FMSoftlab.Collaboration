using MailKit;
using MimeKit;
using MimeKit.Text;
using MimeKit.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
/*using MessagingToolkit.QRCode;
using MessagingToolkit.QRCode.Codec;
using MessagingToolkit.QRCode.Codec.Data;*/
using System.Drawing.Imaging;
using MimeKit.Tnef;

namespace FMSoftlab.Collaboration.Mail.MailKit
{


    public class ReplyVisitor : MimeVisitor
    {
        private string Replytext { get; }
        public string VisitorHtmlBody { get; set; }
        readonly Stack<Multipart> stack = new Stack<Multipart>();
        MimeMessage message;
        MimeEntity body;
        private readonly List<MimeEntity> attachments = new List<MimeEntity>();
        private readonly List<MimePart> embedded = new List<MimePart>();
        /// <summary>
        /// Creates a new ReplyVisitor.
        /// </summary>
        public ReplyVisitor(string replytext)
        {
            Replytext = replytext;
        }

        /// <summary>
        /// Gets the reply.
        /// </summary>
        /// <value>The reply.</value>
        public MimeEntity Body
        {
            get { return body; }
        }

        void Push(MimeEntity entity)
        {
            var multipart = entity as Multipart;

            if (body == null)
            {
                body = entity;
            }
            else
            {
                var parent = stack.Peek();
                parent.Add(entity);
            }

            if (multipart != null)
                stack.Push(multipart);
        }

        void Pop()
        {
            stack.Pop();
        }

        public static string GetOnDateSenderWrote(MimeMessage message)
        {
            var sender = message.Sender != null ? message.Sender : message.From.Mailboxes.FirstOrDefault();
            var name = sender != null ? (!string.IsNullOrEmpty(sender.Name) ? sender.Name : sender.Address) : "someone";

            return string.Format("On {0}, {1} wrote:", message.Date.ToString("f"), name);
        }

        /// <summary>
        /// Visit the specified message.
        /// </summary>
        /// <param name="message">The message.</param>
        public override void Visit(MimeMessage message)
        {
            this.message = message;
            stack.Clear();

            base.Visit(message);
        }

        protected override void VisitMultipartAlternative(MultipartAlternative alternative)
        {
            var multipart = new MultipartAlternative();

            Push(multipart);

            for (int i = 0; i < alternative.Count; i++)
                alternative[i].Accept(this);

            Pop();
        }

        protected override void VisitMultipartRelated(MultipartRelated related)
        {
            var multipart = new MultipartRelated();
            var root = related.Root;

            Push(multipart);

            root.Accept(this);

            for (int i = 0; i < related.Count; i++)
            {
                if (related[i] != root)
                    related[i].Accept(this);
            }

            Pop();
        }

        protected override void VisitMultipart(Multipart multipart)
        {
            foreach (var part in multipart)
            {
                if (part is MultipartAlternative)
                    part.Accept(this);
                else if (part is MultipartRelated)
                    part.Accept(this);
                else if (part is TextPart)
                    part.Accept(this);
            }
        }

        /*bool TryGetImage(string url, out MimePart image)
        {
            UriKind kind;
            int index;
            Uri uri;

            if (Uri.IsWellFormedUriString(url, UriKind.Absolute))
                kind = UriKind.Absolute;
            else if (Uri.IsWellFormedUriString(url, UriKind.Relative))
                kind = UriKind.Relative;
            else
                kind = UriKind.RelativeOrAbsolute;

            try
            {
                uri = new Uri(url, kind);
            }
            catch
            {
                image = null;
                return false;
            }

            for (int i = stack.Count - 1; i >= 0; i--)
            {
                if ((index = stack[i].IndexOf(uri)) == -1)
                    continue;

                image = stack[i][index] as MimePart;
                return image != null;
            }

            image = null;

            return false;
        }
        */

        void HtmlTagCallback(HtmlTagContext ctx, HtmlWriter htmlWriter)
        {
            if (ctx.TagId == HtmlTagId.Body && !ctx.IsEmptyElementTag)
            {
                if (ctx.IsEndTag)
                {
                    // end our opening <blockquote>
                    htmlWriter.WriteEndTag(HtmlTagId.BlockQuote);

                    // pass the </body> tag through to the output
                    ctx.WriteTag(htmlWriter, true);
                }
                else
                {
                    // pass the <body> tag through to the output
                    ctx.WriteTag(htmlWriter, true);                   

                    // prepend the HTML reply with "On {DATE}, {SENDER} wrote:"
                    htmlWriter.WriteStartTag(HtmlTagId.P);
                    
                    string s = GetOnDateSenderWrote(message);
                    if (!string.IsNullOrWhiteSpace(Replytext))
                    {
                        s = $"{Replytext}{Environment.NewLine}{s}";
                    }
                    htmlWriter.WriteText(s);
                    htmlWriter.WriteEndTag(HtmlTagId.P);

                    // Wrap the original content in a <blockquote>
                    htmlWriter.WriteStartTag(HtmlTagId.BlockQuote);
                    htmlWriter.WriteAttribute(HtmlAttributeId.Style, "border-left: 1px #ccc solid; margin: 0 0 0 .8ex; padding-left: 1ex;");

                    ctx.InvokeCallbackForEndTag = true;
                }
            }
            else
            {
                // pass the tag through to the output
                ctx.WriteTag(htmlWriter, true);
            }
        }

        string QuoteText(string text)
        {
            using (var quoted = new StringWriter())
            {
                string s = GetOnDateSenderWrote(message);
                quoted.WriteLine(s);

                using (var reader = new StringReader(text))
                {
                    string line;

                    while ((line = reader.ReadLine()) != null)
                    {
                        quoted.Write("> ");
                        quoted.WriteLine(line);
                    }
                }

                return quoted.ToString();
            }
        }

        protected override void VisitTextPart(TextPart entity)
        {
            string text;

            if (entity.IsHtml)
            {
                var converter = new HtmlToHtml
                {
                    HtmlTagCallback = HtmlTagCallback
                };

                text = converter.Convert(entity.Text);
                VisitorHtmlBody = text;
                Console.WriteLine(text);
            }
            else if (entity.IsFlowed)
            {
                var converter = new FlowedToText();

                text = converter.Convert(entity.Text);
                text = QuoteText(text);
            }
            else
            {
                // quote the original message text
                text = QuoteText(entity.Text);
            }

            var part = new TextPart(entity.ContentType.MediaSubtype.ToLowerInvariant())
            {
                Text = text
            };

            Push(part);
        }
        protected override void VisitTnefPart(TnefPart entity)
        {
            // extract any attachments in the MS-TNEF part
            attachments.AddRange(entity.ExtractAttachments());
        }

        protected override void VisitMessagePart(MessagePart entity)
        {
            // treat message/rfc822 parts as attachments
            attachments.Add(entity);
        }

        protected override void VisitMimePart(MimePart entity)
        {
            // realistically, if we've gotten this far, then we can treat
            // this as an attachment even if the IsAttachment property is
            // false.
            attachments.Add(entity);
        }
    }
    public static class MailHelpers
    {
        /*static Bitmap GetQRCode(int caseId)
        {
            QRCodeEncoder enc = new QRCodeEncoder();
            Bitmap bmp = enc.Encode($"CaseId:{caseId}");
            bmp.Save("pandekths.png", ImageFormat.Png);
            Console.WriteLine($"Width:{bmp.Width}, Height:{bmp.Height}");
            return bmp;
        }
        static void GetQRCode(int caseId, MemoryStream str)
        {
            if (str != null)
            {
                QRCodeEncoder enc = new QRCodeEncoder();
                Bitmap bmp = enc.Encode($"CaseId:{caseId}");
                bmp.Save(str, ImageFormat.Png);
                Console.WriteLine($"Width:{bmp.Width}, Height:{bmp.Height}");
            }
        }
        public static string ScanMessageForQrCode(MimeMessage message)
        {
            string res = "";
            IEnumerable<MimeEntity> ts = message.BodyParts;
            foreach (var t in ts)
            {
                if (t is MimePart mp)
                {
                    Console.WriteLine($"ContentId:{t.ContentId}, ContentLocation:{t.ContentLocation}, ContentType.Name:{t.ContentType.Name}" +
                        $"MediaType:{t.ContentType.MediaType}, MediaSubtype:{t.ContentType.MediaSubtype}");
                    if (string.Equals(t.ContentType.MediaType, "image")
                        && string.Equals(t.ContentType.MediaSubtype, "png")
                        )
                    {
                        using (MemoryStream mStream = new MemoryStream())
                        {
                            mp.Content.DecodeTo(mStream);
                            Bitmap bmp = new Bitmap(mStream);
                            Console.WriteLine($"Width:{bmp.Width}, Height:{bmp.Height}");
                            if (bmp.Width == bmp.Height)
                            {
                                try
                                {
                                    QRCodeDecoder decoder = new QRCodeDecoder();
                                    QRCodeBitmapImage qbm = new QRCodeBitmapImage(bmp);
                                    res = decoder.Decode(qbm);
                                }
                                catch
                                {
                                    res = "";
                                }
                                Console.WriteLine(res);
                            }
                        }
                    }
                }
            }
            return res;
        }*/

        public static MimeMessage ReplyHtml(MimeMessage message, MailboxAddress from, bool replyToAll,string replytext, FmSoftlabBaseMailKitAttachment attachment)
        {
            AttachmentCollection LinkedResources = new AttachmentCollection();
            var visitor = new ReplyVisitor("");
            var reply = new MimeMessage();

            reply.From.Add(from);

            // reply to the sender of the message
            if (message.ReplyTo.Count > 0)
            {
                reply.To.AddRange(message.ReplyTo);
            }
            else if (message.From.Count > 0)
            {
                reply.To.AddRange(message.From);
            }
            else if (message.Sender != null)
            {
                reply.To.Add(message.Sender);
            }

            if (replyToAll)
            {
                // include all of the other original recipients (removing ourselves from the list)
                reply.To.AddRange(message.To.Mailboxes.Where(x => x.Address != from.Address));
                reply.Cc.AddRange(message.Cc.Mailboxes.Where(x => x.Address != from.Address));
            }

            // set the reply subject
            if (!message.Subject.StartsWith("Re:", StringComparison.OrdinalIgnoreCase))
                reply.Subject = "Re: " + message.Subject;
            else
                reply.Subject = message.Subject;

            // construct the In-Reply-To and References headers
            if (!string.IsNullOrEmpty(message.MessageId))
            {
                reply.InReplyTo = message.MessageId;
                foreach (var id in message.References)
                    reply.References.Add(id);
                reply.References.Add(message.MessageId);
            }

            visitor.Visit(message);
            reply.Body = visitor.Body ?? new TextPart("plain") { Text = ReplyVisitor.GetOnDateSenderWrote(message) + Environment.NewLine };

            /*BodyBuilder bodyBuilder = new BodyBuilder();
            var image = bodyBuilder.LinkedResources.Add("pandekths.png");
            image.ContentId = MimeUtils.GenerateMessageId();
            bodyBuilder.HtmlBody = string.Format(@"<p>Hey!</p><img src=""cid:{0}"">", image.ContentId);*/
            //MimePart att = GetAttachmentWithQrEncoding("test1", 123456);
            //LinkedResources.Add(att);
            string htmlbody = $"<p>{replytext}</p><img src=\"cid:{attachment.Attachment.ContentId}\">"+visitor.VisitorHtmlBody;
            var text = new TextPart("html");
            text.ContentId = MimeUtils.GenerateMessageId();
            text.Text = htmlbody;
            var related = new MultipartRelated
            {
                Root = text
            };
            related.Add(attachment.Attachment);
            reply.Body = related;
            return reply;
        }

        public static MimeMessage ReplyHtml(MimeMessage message, MailboxAddress from, bool replyToAll, string replytext)
        {
            AttachmentCollection LinkedResources = new AttachmentCollection();
            var visitor = new ReplyVisitor(replytext);
            var reply = new MimeMessage();

            reply.From.Add(from);

            // reply to the sender of the message
            if (message.ReplyTo.Count > 0)
            {
                reply.To.AddRange(message.ReplyTo);
            }
            else if (message.From.Count > 0)
            {
                reply.To.AddRange(message.From);
            }
            else if (message.Sender != null)
            {
                reply.To.Add(message.Sender);
            }

            if (replyToAll)
            {
                // include all of the other original recipients (removing ourselves from the list)
                reply.To.AddRange(message.To.Mailboxes.Where(x => x.Address != from.Address));
                reply.Cc.AddRange(message.Cc.Mailboxes.Where(x => x.Address != from.Address));
            }

            // set the reply subject
            if (!message.Subject.StartsWith("Re:", StringComparison.OrdinalIgnoreCase))
                reply.Subject = "Re: " + message.Subject;
            else
                reply.Subject = message.Subject;

            // construct the In-Reply-To and References headers
            if (!string.IsNullOrEmpty(message.MessageId))
            {
                reply.InReplyTo = message.MessageId;
                foreach (var id in message.References)
                    reply.References.Add(id);
                reply.References.Add(message.MessageId);
            }

            visitor.Visit(message);
            reply.Body = visitor.Body ?? new TextPart("plain") { Text = ReplyVisitor.GetOnDateSenderWrote(message) + Environment.NewLine };

            return reply;
        }

        public static MimeMessage Forward(MimeMessage original, MailboxAddress from, IEnumerable<InternetAddress> to)
        {
            var message = new MimeMessage();
            message.From.Add(from);
            message.To.AddRange(to);

            // set the forwarded subject
            if (!original.Subject.StartsWith("FW:", StringComparison.OrdinalIgnoreCase))
                message.Subject = "FW: " + original.Subject;
            else
                message.Subject = original.Subject;

            // quote the original message text
            using (var text = new StringWriter())
            {
                text.WriteLine();
                text.WriteLine("-----Original Message-----");
                text.WriteLine("From: {0}", original.From);
                text.WriteLine("Sent: {0}", DateUtils.FormatDate(original.Date));
                text.WriteLine("To: {0}", original.To);
                text.WriteLine("Subject: {0}", original.Subject);
                text.WriteLine();

                text.Write(original.TextBody);

                message.Body = new TextPart("plain")
                {
                    Text = text.ToString()
                };
            }
            return message;
        }

        public static MimeMessage Reply2(MimeMessage message, MailboxAddress from, bool replyToAll)
        {
            var reply = new MimeMessage();

            reply.From.Add(from);

            // reply to the sender of the message
            if (message.ReplyTo.Count > 0)
            {
                reply.To.AddRange(message.ReplyTo);
            }
            else if (message.From.Count > 0)
            {
                reply.To.AddRange(message.From);
            }
            else if (message.Sender != null)
            {
                reply.To.Add(message.Sender);
            }

            if (replyToAll)
            {
                // include all of the other original recipients (removing ourselves from the list)
                reply.To.AddRange(message.To.Mailboxes.Where(x => x.Address != from.Address));
                reply.Cc.AddRange(message.Cc.Mailboxes.Where(x => x.Address != from.Address));
            }

            // set the reply subject
            if (!message.Subject.StartsWith("Re:", StringComparison.OrdinalIgnoreCase))
                reply.Subject = "Re: " + message.Subject;
            else
                reply.Subject = message.Subject;

            // construct the In-Reply-To and References headers
            if (!string.IsNullOrEmpty(message.MessageId))
            {
                reply.InReplyTo = message.MessageId;
                foreach (var id in message.References)
                    reply.References.Add(id);
                reply.References.Add(message.MessageId);
            }

            // quote the original message text
            using (var quoted = new StringWriter())
            {
                var sender = message.Sender ?? message.From.Mailboxes.FirstOrDefault();
                var name = sender != null ? (!string.IsNullOrEmpty(sender.Name) ? sender.Name : sender.Address) : "someone";

                quoted.WriteLine("On {0}, {1} wrote:", message.Date.ToString("f"), name);
                using (var reader = new StringReader(message.TextBody))
                {
                    string line;

                    while ((line = reader.ReadLine()) != null)
                    {
                        quoted.Write("> ");
                        quoted.WriteLine(line);
                    }
                }
                
                reply.Body = new TextPart("plain")
                {
                    Text = quoted.ToString()
                };
            }
            return reply;
        }
    }
}
