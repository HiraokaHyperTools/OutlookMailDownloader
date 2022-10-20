using Azure.Core;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Azure.Identity;
using OutlookMailDownloader.Models;

namespace OutlookMailDownloader.Utils
{
    internal class MailReader
    {
        private static readonly string TOKEN_CACHE_NAME = "OutlookMailDownloader";

        /// <summary>
        /// https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Authentication/appId/1077436d-94d8-4938-8d4a-d2a3221df47f/objectId/86146439-d833-494e-872d-176b20a0088d/isMSAApp~/false/defaultBlade/Overview/appSignInAudience/AzureADandPersonalMicrosoftAccount/servicePrincipalCreated~/true
        /// </summary>
        private static readonly string clientId = "1077436d-94d8-4938-8d4a-d2a3221df47f";

        public async Task ReceiveAsync(
            string authFile,
            Func<string, string, DateTimeOffset?, byte[], Task> saveAttachment,
            Func<string, Task<bool>> needDownload,
            CancellationToken cancellationToken
        )
        {
            // https://docs.microsoft.com/ja-jp/graph/sdks/create-client?tabs=CS
            var scopes = new[] { "User.Read", "Mail.Read", };

            // Multi-tenant apps can use "common",
            // single-tenant apps must use the tenant ID from the Azure portal
            var tenantId = "common";

            // Value from app registration

            // using Azure.Identity;
            var options = new InteractiveBrowserCredentialOptions
            {
                TenantId = tenantId,
                ClientId = clientId,
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                // MUST be http://localhost or http://localhost:PORT
                // See https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/System-Browser-on-.Net-Core
                RedirectUri = new Uri("http://localhost"),
                LoginHint = "",
                TokenCachePersistenceOptions = new TokenCachePersistenceOptions
                {
                    Name = TOKEN_CACHE_NAME
                },
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.interactivebrowsercredential
            TokenCredential credential;
            InteractiveBrowserCredential browserCredential;
            AuthenticationRecord authRecord;

            // https://docs.microsoft.com/en-us/dotnet/api/azure.identity.tokencachepersistenceoptions?view=azure-dotnet
            // Check if an AuthenticationRecord exists on disk.
            // If it does not exist, get one and serialize it to disk.
            // If it does exist, load it from disk and deserialize it.
            if (!System.IO.File.Exists(authFile))
            {
                // Construct a credential with TokenCachePersistenceOptions specified to ensure that the token cache is persisted to disk.
                // We can also optionally specify a name for the cache to avoid having it cleared by other applications.
                browserCredential = new InteractiveBrowserCredential(options);

                // Call AuthenticateAsync to fetch a new AuthenticationRecord.
                authRecord = await browserCredential.AuthenticateAsync(
                    new TokenRequestContext(scopes, null, null, tenantId),
                    cancellationToken
                );

                // Serialize the AuthenticationRecord to disk so that it can be re-used across executions of this initialization code.
                using (var authRecordStream = new FileStream(authFile, FileMode.Create, FileAccess.Write))
                {
                    await authRecord.SerializeAsync(authRecordStream);
                }

                credential = browserCredential;
            }
            else
            {
                // Load the previously serialized AuthenticationRecord from disk and deserialize it.
                using (var authRecordStream = new FileStream(authFile, FileMode.Open, FileAccess.Read))
                {
                    authRecord = await AuthenticationRecord.DeserializeAsync(authRecordStream);

                    // Construct a new client with our TokenCachePersistenceOptions with the addition of the AuthenticationRecord property.
                    // This tells the credential to use the same token cache in addition to which account to try and fetch from cache when GetToken is called.
                    options.AuthenticationRecord = authRecord;
                    browserCredential = new InteractiveBrowserCredential(options);

                    //await credential.AuthenticateAsync(
                    //    new TokenRequestContext(scopes, null, null, tenantId),
                    //    cancellationToken
                    //);

                    credential = browserCredential;
                }
            }

            //Console.WriteLine(credential.GetToken(new TokenRequestContext(scopes)).Token);

            var graphClient = new GraphServiceClient(credential, scopes);

            var msgs = await graphClient.Me
                .Messages
                .Request()
                .Expand("attachments")
                .GetAsync();

            while (true)
            {
                if (msgs.CurrentPage != null)
                {
                    foreach (var msg in msgs.CurrentPage)
                    {
                        if (!(await needDownload(msg.Id)))
                        {
                            return;
                        }

                        Console.WriteLine($"{msg.LastModifiedDateTime} {msg.Subject} ({msg.Attachments?.Count})");

                        if (msg.HasAttachments == true)
                        {
                            var attachments = msg.Attachments;

                            while (true)
                            {
                                if (attachments?.CurrentPage != null)
                                {
                                    foreach (var attachment in attachments.CurrentPage.OfType<FileAttachment>())
                                    {
                                        Console.WriteLine($"- {attachment.Name} {attachment.ContentBytes?.Length}");
                                        if (attachment.ContentBytes != null)
                                        {
                                            await saveAttachment(
                                                msg.Subject,
                                                attachment.Name,
                                                attachment.LastModifiedDateTime,
                                                attachment.ContentBytes
                                            );
                                        }
                                    }
                                }

                                if (attachments?.NextPageRequest == null)
                                {
                                    break;
                                }

                                attachments = await attachments.NextPageRequest.GetAsync();
                            }
                        }
                    }
                }

                if (msgs.NextPageRequest == null)
                {
                    return;
                }

                msgs = await msgs.NextPageRequest.GetAsync();
            }
        }
    }
}
