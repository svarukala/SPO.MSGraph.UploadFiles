using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using MicrosoftGraphFilesUpload.TokenStorage;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.IO;
using System;

namespace MicrosoftGraphFilesUpload.Helpers
{
    public static class GraphHelper
    {
        // Load configuration settings from PrivateSettings.config
        private static string appId = ConfigurationManager.AppSettings["ida:AppId"];
        private static string appSecret = ConfigurationManager.AppSettings["ida:AppSecret"];
        private static string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];
        private static string graphScopes = ConfigurationManager.AppSettings["ida:AppScopes"];
        private static string hostname = ConfigurationManager.AppSettings["ida:HostName"];

        public static async Task<IEnumerable<Event>> GetEventsAsync()
        {
            var graphClient = GetAuthenticatedClient();

            var events = await graphClient.Me.Events.Request()
                .Select("subject,organizer,start,end")
                .OrderBy("createdDateTime DESC")
                .GetAsync();

            return events.CurrentPage;
        }


        public static async Task<DriveItem> UploadFileAsync(Stream fileStream, string fileName, bool isLargeFile, string siteUrl)
        {
            var graphClient = GetAuthenticatedClient();
            DriveItem uploadedFile = null;
            siteUrl = string.IsNullOrEmpty(siteUrl.Trim())? "/sites/web01" : siteUrl;
            //var hostname = "m365x130314.sharepoint.com";

            try
            {
                var site = await graphClient.Sites.GetByPath(siteUrl, hostname).Request().GetAsync();
                //var lists = await graphClient.Sites.GetByPath(siteUrl, hostname).Lists.Request().GetAsync();
                //var listId = lists.First(p => p.Name == "Documents").Id;

                if (!isLargeFile)
                    uploadedFile = await graphClient.Sites[site.Id].Drive.Root.ItemWithPath(fileName).Content.Request().PutAsync<DriveItem>(fileStream);
                //uploadedFile = await graphClient.Sites.GetByPath(siteUrl, hostname).Lists[listId].Drive.Root.ItemWithPath(fileName).Content.Request().PutAsync<DriveItem>(fileStream);
                //uploadedFile = await graphClient.Sites[hostname + siteUrl].Drive.Root.ItemWithPath(fileName).Content.Request().PutAsync<DriveItem>(fileStream);
                else
                {
                    UploadSession uploadSession = null;
                    //uploadSession = await graphClient.Sites["root"].SiteWithPath("/sites/team01").Drive.Root.ItemWithPath(fileName).CreateUploadSession().Request().PostAsync();
                    uploadSession = await graphClient.Sites[site.Id].Drive.Root.ItemWithPath(fileName).CreateUploadSession().Request().PostAsync();

                    if (uploadSession != null)
                    {
                        // Chunk size must be divisible by 320KiB, our chunk size will be slightly more than 1MB
                        int maxSizeChunk = (320 * 1024) * 16;
                        ChunkedUploadProvider uploadProvider = new ChunkedUploadProvider(uploadSession, graphClient, fileStream, maxSizeChunk);
                        var chunkRequests = uploadProvider.GetUploadChunkRequests();
                        var exceptions = new List<Exception>();
                        var readBuffer = new byte[maxSizeChunk];
                        foreach (var request in chunkRequests)
                        {
                            var result = await uploadProvider.GetChunkRequestResponseAsync(request, readBuffer, exceptions);

                            if (result.UploadSucceeded)
                            {
                                uploadedFile = result.ItemResponse;
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                throw ex;
                
            }

            return uploadedFile;
        }


        private static GraphServiceClient GetAuthenticatedClient()
        {
            return new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        var idClient = ConfidentialClientApplicationBuilder.Create(appId)
                            .WithRedirectUri(redirectUri)
                            .WithClientSecret(appSecret)
                            .Build();
                        var tokenStore = new SessionTokenStore(idClient.UserTokenCache,
                                HttpContext.Current, ClaimsPrincipal.Current);

                        var accounts = await idClient.GetAccountsAsync();

                        // By calling this here, the token can be refreshed
                        // if it's expired right before the Graph call is made
                        var scopes = graphScopes.Split(' ');
                        var result = await idClient.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                            .ExecuteAsync();

                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    }));
        }


        public static async Task<User> GetUserDetailsAsync(string accessToken)
        {
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", accessToken);
                    }));

            return await graphClient.Me.Request().GetAsync();
        }
    }
}