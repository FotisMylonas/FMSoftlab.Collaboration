using Azure.Communication.Email;
using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Abstractions;
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace FMSoftlab.Collaboration.Mail.AzureEmailClient
{
    public class AzureRestEmailClient
    {
        private readonly ILogger<AzureRestEmailClient> _log;
        private readonly IMailServerConnectionSettings _mailServerConnectionSettings;
        public AzureRestEmailClient(IMailServerConnectionSettings mailServerConnectionSettings, ILogger<AzureRestEmailClient> log)
        {
            _log=log;
            _mailServerConnectionSettings=mailServerConnectionSettings;
        }
        public async Task<SendMailResult> SendAsync(ISendMailAttributes email)
        {
            SendMailResult res = SendMailResult.Failed("invalid request");
            if (_mailServerConnectionSettings is null)
            {
                _log?.LogWarning("AzureRestEmailClient, no settings");
                return res;
            }
            _log?.LogDebug($"AzureRestEmailClient, Sending email using azure, uri:{_mailServerConnectionSettings.Server}");
            if (email is null)
            {
                _log?.LogWarning("AzureRestEmailClient, no email attributes");
                return res;
            }
            if (email.ToAddresses.Count<=0)
            {
                _log?.LogWarning("AzureRestEmailClient, no recipients");
                return res;
            }
            if (string.IsNullOrWhiteSpace(email.Subject) && string.IsNullOrWhiteSpace(email.Body))
            {
                _log?.LogWarning("AzureRestEmailClient, no subject / body");
                return res;
            }
            //DefaultAzureCredentialOptions defaultAzureCredentialOptions = new DefaultAzureCredentialOptions() { };
            EmailClient emailClient = new EmailClient(new Uri(_mailServerConnectionSettings.Server), new DefaultAzureCredential());
            EmailRecipients recipients = new EmailRecipients();
            foreach (FMEmailAddress a in email.ToAddresses)
            {
                recipients.To.Add(new EmailAddress(a.Address, a.Name));
            }
            foreach (FMEmailAddress a in email.CcAddresses)
            {
                recipients.CC.Add(new EmailAddress(a.Address, a.Name));
            }
            EmailContent content = new EmailContent(email.Subject) { Html=email.Body };
            EmailMessage emailMessage = new EmailMessage(email.FromAddress, recipients, content);
            try
            {
                string recep = string.Join(",", emailMessage.Recipients.To.Select(s => s.Address).ToList());
                _log?.LogDebug($"AzureRestEmailClient, will send email, uri:{_mailServerConnectionSettings.Server}, Subject:{emailMessage.Content.Subject} recepients:{recep}");
                EmailSendOperation emailSendOperation = await emailClient.SendAsync(Azure.WaitUntil.Completed, emailMessage);
                string rawResponse = string.Empty;
                var response = emailSendOperation.GetRawResponse();
                using (var reader = new StreamReader(response.ContentStream))
                {
                    rawResponse = await reader.ReadToEndAsync();
                }
                EmailSendResult sendResult = emailSendOperation.Value;
                _log?.LogInformation(
                    "AzureRestEmailClient: Email send operation completed. Id:{OperationId}, Send Result Status:{SendResultStatus}, Reason:{ReasonPhrase}, HttpStatus:{HttpStatusCode}, IsError:{IsError}, Uri:{Uri}, Subject:{Subject}, Recipients:{Recipients}, RawResponse:{RawResponse}",
                    emailSendOperation.Id,
                    sendResult.Status,
                    response.ReasonPhrase,
                    response.Status,
                    response.IsError,
                    _mailServerConnectionSettings.Server,
                    emailMessage.Content.Subject,
                    recep,
                    rawResponse
                );
                if (response.IsError)
                {
                    return SendMailResult.Failed(response.ReasonPhrase);
                }
                else return SendMailResult.SuccessNoId;
            }
            catch (Exception ex)
            {
                string message = ex.Message;
                _log?.LogError(message);
                return SendMailResult.Failed(message);
            }
        }
    }
}