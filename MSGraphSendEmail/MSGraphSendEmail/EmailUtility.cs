
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace MSGraphSendEmail
{
    public class EmailUtility
    {
        private readonly IConfidentialClientApplication _confidentialClientApplication;
        private readonly EmailSettings _emailSettings;
        private static AccessToken? _accessToken;

        static EmailUtility()
        {
            _accessToken = null!;
        }

        public EmailUtility(IConfidentialClientApplication confidentialClientApplication,
            IOptions<EmailSettings> emailSettings
            )
        {
            _confidentialClientApplication = confidentialClientApplication;
            _emailSettings = emailSettings.Value;

        }

        /// <summary>
        /// Calls Send Email Mehtod which calls the MS Graph using an authenticated Http client
        /// </summary>

        public async Task<bool> SendMail(string recipient, string bcc, string cc, string subject, string body, string attachment)
        {
            // First We Get the JWT Token 
            var accessToken = await GetToken();

            if (accessToken != null && !string.IsNullOrEmpty(accessToken.Token))
            {
                Message message = new Message
                {
                    Subject = subject,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = body
                    }
                };

                List<Recipient> toRecipients = new List<Recipient>();
                List<Recipient> ccRecipients = new List<Recipient>();
                List<Recipient> bccRecipients = new List<Recipient>();

                // Set the recipient address of the mail message             
                if (!string.IsNullOrEmpty(recipient))
                {
                    string[] strRecipient = recipient.Replace(";", ",").TrimEnd(',').Split(new char[] { ',' });

                    // Set the Bcc address of the mail message 
                    for (int intCount = 0; intCount < strRecipient.Length; intCount++)
                    {
                        var emailRecipient = AddEmailAddress(recipient);
                        toRecipients.Add(emailRecipient);
                    }
                }
                // Check if the bcc value is nothing or an empty string 
                if (!string.IsNullOrEmpty(bcc))
                {
                    string[] strBCC = bcc.Split(new char[] { ',' });

                    // Set the Bcc address of the mail message 
                    for (int intCount = 0; intCount < strBCC.Length; intCount++)
                    {
                        var emailRecipient = AddEmailAddress(recipient);
                        bccRecipients.Add(emailRecipient);
                    }
                }

                // Check if the cc value is nothing or an empty value 
                if (!string.IsNullOrEmpty(cc))
                {
                    // Set the CC address of the mail message 
                    string[] strCC = cc.Split(new char[] { ',' });
                    for (int intCount = 0; intCount < strCC.Length; intCount++)
                    {
                        var emailRecipient = AddEmailAddress(recipient);
                        ccRecipients.Add(emailRecipient);
                    }
                }

                #region --> attachment
                if (!string.IsNullOrEmpty(attachment))
                {
                    // Create the message with attachment.
                    byte[] contentBytes = System.IO.File.ReadAllBytes(string.Empty);
                    string contentType = "image/png";
                    MessageAttachmentsCollectionPage attachments = new MessageAttachmentsCollectionPage();
                    attachments.Add(new FileAttachment
                    {
                        ODataType = "#microsoft.graph.fileAttachment",
                        ContentBytes = contentBytes,
                        ContentType = contentType,
                        ContentId = "testing",
                        Name = "testing.png"
                    });
                    message.Attachments = attachments;
                }
                #endregion

                message.ToRecipients = toRecipients;
                message.CcRecipients = null;
                message.BccRecipients = null;

                //Create GraphServiceClient and add the access token in request header.
                GraphServiceClient graphServiceClient =
                    new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                    {
                        // Add the access token in the Authorization header of the API request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken.Token);
                    })
                );

                //Call the SendMail API of graph service client 
                await graphServiceClient.Users[_emailSettings.FromEmail]
                         .SendMail(message, false)
                         .Request()
                         .PostAsync();

                return true;

            }
            else
            {
                return false;
            }
        }

        public Recipient AddEmailAddress(string address)
        {
            Recipient recipient = new Recipient();
            EmailAddress emailAddress = new EmailAddress();
            emailAddress.Address = address;
            recipient.EmailAddress = emailAddress;
            return recipient;
        }

        public async Task<AccessToken> GetToken()
        {
            if (_accessToken is { Expired: false })
            {
                return _accessToken;
            }
            _accessToken = await FetchToken();
            return _accessToken;
        }

        private async Task<AccessToken> FetchToken()
        {
            string[] scopes = new string[] { $"{_emailSettings.ApiUrl}.default" };
            AuthenticationResult result = await _confidentialClientApplication.AcquireTokenForClient(scopes)
                .ExecuteAsync().ConfigureAwait(false);
            AccessToken accessToken = new AccessToken(result.AccessToken, result.ExpiresOn.DateTime);
            return accessToken;
        }
    }
}
