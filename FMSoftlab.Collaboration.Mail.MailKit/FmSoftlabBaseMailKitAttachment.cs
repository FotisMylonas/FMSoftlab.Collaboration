//using MessagingToolkit.QRCode.Codec;
using MimeKit;
using MimeKit.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FMSoftlab.Collaboration.Mail.MailKit
{
    public abstract class FmSoftlabBaseMailKitAttachment : IDisposable
    {
        bool disposed = false;
        protected string Filename { get; }
        protected MemoryStream Stream { get; }
        /*protected string MediaType { get; }
        protected string MediaSubtype { get; }*/
        protected abstract void LoadMemoryStream(MemoryStream Stream);
        private void DoLoadMemoryStream()
        {
            LoadMemoryStream(Stream);
        }
        private MimePart mAttachment = null;
        public MimePart Attachment
        {
            get
            {
                if (mAttachment == null)
                {
                    mAttachment = BuildAttachment();
                }
                return mAttachment;
            }
        }
        private MimePart BuildAttachment()
        {
            MimePart res = null;
            Stream.Position = 0;
            DoLoadMemoryStream();
            Stream.Position = 0;
            string mimetype = MimeTypes.GetMimeType(Filename);
            res = new MimePart(mimetype)
            {
                Content = new MimeContent(Stream, ContentEncoding.Default),
                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                ContentTransferEncoding = ContentEncoding.Base64,
                FileName = this.Filename,
                ContentId = MimeUtils.GenerateMessageId()
            };
            return res;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // Protected implementation of Dispose pattern.
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                try
                {
                    try
                    {
                        Stream.Close();
                    }
                    finally
                    {
                        ((IDisposable)Stream).Dispose();
                        GC.SuppressFinalize(this);
                    }
                }
                catch
                {

                }
            }
            disposed = true;
        }

        public FmSoftlabBaseMailKitAttachment(string filename)
        {
            Stream = new MemoryStream();
            /*MediaType = mediaType;
            MediaSubtype = mediaSubtype;*/
            Filename = filename;
        }
    }

    public class FmSoftlabCalendarAttachment : FmSoftlabBaseMailKitAttachment
    {
        private FileAttachment Fa { get; }
        public FmSoftlabCalendarAttachment(FileAttachment fa) : base(fa.OriginalFilename)
        {
            Fa = fa;
        }
        protected override void LoadMemoryStream(MemoryStream stream)
        {
            stream.Write(Fa.Content, 0, Fa.Content.Length);
        }
    }
            


    public static class MailKitAttachmentFactory
    {
        public static MimePart GetAttachment(FileAttachment attachment)
        {
            MimePart res = null;
            using (MemoryStream ms = new MemoryStream(attachment.Content))
            {
                string mimeType = MimeTypes.GetMimeType(attachment.OriginalFilename);
                res = new MimePart(mimeType)
                {
                    Content = new MimeContent(ms, ContentEncoding.Default),
                    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                    ContentTransferEncoding = ContentEncoding.Base64,
                    FileName = Path.GetFileName(attachment.OriginalFilename),
                    ContentId = MimeUtils.GenerateMessageId()
                };
            }
            return res;
        }
    }

    public abstract class FmSoftlabMailKitPngAttachment : FmSoftlabBaseMailKitAttachment
    {
        protected abstract Bitmap GetBitmap();
        public FmSoftlabMailKitPngAttachment(string filename) : base(filename)
        {

        }
        protected override void LoadMemoryStream(MemoryStream Stream)
        {
            Bitmap bmap = GetBitmap();
            if (bmap != null)
            {
                bmap.Save(Stream, ImageFormat.Png);
            }
        }
    }

    /*public class FmSoftlabMailQrEncodeAttachment : FmSoftlabMailKitPngAttachment
    {
        public int CaseId { get; }
        protected override Bitmap GetBitmap()
        {
            QRCodeEncoder enc = new QRCodeEncoder();
            return enc.Encode($"CaseId:{CaseId}");
        }
        public FmSoftlabMailQrEncodeAttachment(int caseId) : base($"pandekthspng{caseId}.png")
        {
            CaseId = caseId;
        }
    }*/

}
