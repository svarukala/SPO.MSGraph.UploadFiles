// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System.Collections;
using System.Collections.Generic;
using System.Security.Claims;
using System.Threading;
using System.Web;

namespace MicrosoftGraphFilesUpload.TokenStorage
{
    // Simple class to serialize into the session
    public class CachedUser
    {
        public string DisplayName { get; set; }
        public string Email { get; set; }
        public string Avatar { get; set; }
    }

    public class SessionTokenStore
    {
        private static readonly ReaderWriterLockSlim sessionLock = new ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion);

        private static HttpContext httpContext = null;
        private static HttpContextBase httpContextBase = null;
        private string tokenCacheKey = string.Empty;
        private string userCacheKey = string.Empty;
        private static Hashtable tokenCacheTable = new Hashtable();
        public SessionTokenStore(ITokenCache tokenCache, HttpContext context, ClaimsPrincipal user, HttpContextBase contextBase)
        //public SessionTokenStore(ITokenCache tokenCache, HttpContextBase context, ClaimsPrincipal user)
        {

            httpContext = context;
            httpContextBase = contextBase;

            if (tokenCache != null)
            {
                tokenCache.SetBeforeAccess(BeforeAccessNotification);
                tokenCache.SetAfterAccess(AfterAccessNotification);
            }

            var userId = GetUsersUniqueId(user);
            tokenCacheKey = $"{userId}_TokenCache";
            userCacheKey = $"{userId}_UserCache";

            //tokenCacheTable = new Hashtable();
        }

        public SessionTokenStore(ITokenCache tokenCache, HttpContext context, ClaimsPrincipal user)
        //public SessionTokenStore(ITokenCache tokenCache, HttpContextBase context, ClaimsPrincipal user)
        {

            httpContext = context;

            if (tokenCache != null)
            {
                tokenCache.SetBeforeAccess(BeforeAccessNotification);
                tokenCache.SetAfterAccess(AfterAccessNotification);
            }

            var userId = GetUsersUniqueId(user);
            tokenCacheKey = $"{userId}_TokenCache";
            userCacheKey = $"{userId}_UserCache";
        }

        public bool HasData()
        {
            return (httpContext.Session[tokenCacheKey] != null &&
                ((byte[])httpContext.Session[tokenCacheKey]).Length > 0);
        }

        public void Clear()
        {
            sessionLock.EnterWriteLock();

            try
            {
                httpContext.Session.Remove(tokenCacheKey);
            }
            finally
            {
                sessionLock.ExitWriteLock();
            }
        }

        private void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            sessionLock.EnterReadLock();

            try
            {
                args.TokenCache.DeserializeMsalV3((byte[])tokenCacheTable[tokenCacheKey]);
                //args.TokenCache.DeserializeMsalV3((byte[])httpContext.Session[tokenCacheKey]);
                /*
                if (httpContextBase != null)
                {
                    // Load the cache from the session
                    args.TokenCache.DeserializeMsalV3((byte[])httpContextBase.Session[tokenCacheKey]);
                }
                else
                {
                    args.TokenCache.DeserializeMsalV3((byte[])httpContext.Session[tokenCacheKey]);
                }
                */
            }
            finally
            {
                sessionLock.ExitReadLock();
            }
        }

        private void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            if (args.HasStateChanged)
            {
                sessionLock.EnterWriteLock();

                try
                {
                    // Store the serialized cache in the session
                    if(!tokenCacheTable.ContainsKey(tokenCacheKey))
                        tokenCacheTable.Add(tokenCacheKey, args.TokenCache.SerializeMsalV3());
                    httpContext.Session[tokenCacheKey] = args.TokenCache.SerializeMsalV3();
                    //httpContextBase.Session[tokenCacheKey] = args.TokenCache.SerializeMsalV3();
                }
                finally
                {
                    sessionLock.ExitWriteLock();
                }
            }
        }

        public void SaveUserDetails(CachedUser user)
        {

            sessionLock.EnterWriteLock();
            httpContext.Session[userCacheKey] = JsonConvert.SerializeObject(user);
            sessionLock.ExitWriteLock();
        }

        public CachedUser GetUserDetails()
        {
            sessionLock.EnterReadLock();
            var cachedUser = JsonConvert.DeserializeObject<CachedUser>((string)httpContext.Session[userCacheKey]);
            sessionLock.ExitReadLock();
            return cachedUser;
        }

        private string GetUsersUniqueId(ClaimsPrincipal user)
        {
            // Combine the user's object ID with their tenant ID

            if (user != null)
            {
                var userObjectId = user.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value ??
                    user.FindFirst("oid").Value;

                var userTenantId = user.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value ??
                    user.FindFirst("tid").Value;

                if (!string.IsNullOrEmpty(userObjectId) && !string.IsNullOrEmpty(userTenantId))
                {
                    return $"{userObjectId}.{userTenantId}";
                }
            }

            return null;
        }
    }
}