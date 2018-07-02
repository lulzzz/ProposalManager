// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

using Microsoft.Extensions.Options;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Identity.Client;
using Microsoft.Identity;
using Microsoft.Graph;
using ApplicationCore;
using ApplicationCore.Interfaces;
using Infrastructure.Identity;
using Infrastructure.Identity.Proxy;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Authentication;
//using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace Infrastructure.GraphApi
{
    /// <summary>
    /// Provider to get the access token 
    /// </summary>
    public class GraphAuthProvider : IGraphAuthProvider
    {
        private readonly IMemoryCache _memoryCache;
        private TokenCache _userTokenCache;

        // Properties used to get and manage an access token.
        private readonly string _clientId;
        private readonly string _aadInstance;
        private readonly ClientCredential _credential;
        private readonly string _appSecret;
        private readonly string[] _scopes;
        private readonly string _redirectUri;
        private readonly string _graphResourceId;
        private readonly string _tenantId;
        private readonly string _authority;
        private readonly DownstreamServiceProxyOptions _serviceOptions;
        private readonly IHttpContextAccessor _httpContextAccessor;


        public GraphAuthProvider(
            IMemoryCache memoryCache, 
            IConfiguration configuration,
            IOptions<DownstreamServiceProxyOptions> serviceOptions,
            IHttpContextAccessor httpContextAccessor)
        {
            var azureOptions = new AzureAdOptions();
            configuration.Bind("AzureAd", azureOptions);

            _clientId = azureOptions.ClientId;
            _aadInstance = azureOptions.Instance;
            _appSecret = azureOptions.ClientSecret;
            _credential = new Microsoft.Identity.Client.ClientCredential(azureOptions.ClientSecret); // For development mode purposes only. Production apps should use a client certificate.
            _scopes = azureOptions.GraphScopes.Split(new[] { ' ' });
            _redirectUri = azureOptions.BaseUrl + azureOptions.CallbackPath;
            _graphResourceId = azureOptions.GraphResourceId;
            _tenantId = azureOptions.TenantId;

            _memoryCache = memoryCache;

            _authority = azureOptions.Authority;
            _serviceOptions = serviceOptions.Value;
            _httpContextAccessor = httpContextAccessor;
        }

        // Gets an access token. First tries to get the access token from the token cache.
        // Using password (secret) to authenticate. Production apps should use a certificate.
        public async Task<string> GetUserAccessTokenAsync(string userId)
        {
            if (_userTokenCache == null) _userTokenCache = new SessionTokenCache(userId, _memoryCache).GetCacheInstance();

            var cca = new ConfidentialClientApplication(
                _clientId,
                _redirectUri,
                _credential,
                _userTokenCache,
                null);

            var originalToken = await _httpContextAccessor.HttpContext.GetTokenAsync("access_token");

            var userAssertion = new UserAssertion(originalToken,
                "urn:ietf:params:oauth:grant-type:jwt-bearer");

            try
            {
                var result = await cca.AcquireTokenOnBehalfOfAsync(_scopes, userAssertion);

                return result.AccessToken;
            }
            catch (Exception ex)
            {
                // Unable to retrieve the access token silently.
                throw new ServiceException(new Error
                {
                    Code = GraphErrorCode.AuthenticationFailure.ToString(),
                    Message = $"Caller needs to authenticate. Unable to retrieve the access token silently. error: {ex}"
                });
            }
        }


        // Gets an access token. First tries to get the access token from the token cache.
        // This app uses a password (secret) to authenticate. Production apps should use a certificate.
        public async Task<string> GetAppAccessTokenAsync()
        {

            try
            {
                var authorityFormat = "https://login.microsoftonline.com/{0}/v2.0";
                ConfidentialClientApplication daemonClient = new ConfidentialClientApplication(_clientId, String.Format(authorityFormat, _tenantId), _redirectUri, _credential, null, new TokenCache());

                var msGraphScope = "https://graph.microsoft.com/.default";
                AuthenticationResult result = await daemonClient.AcquireTokenForClientAsync(new string[] { msGraphScope });

                return result.AccessToken;
            }
            catch (Exception ex)
            {
                // Unable to retrieve the access token silently.
                throw new ServiceException(new Error
                {
                    Code = GraphErrorCode.AuthenticationFailure.ToString(),
                    Message = $"Caller needs to authenticate. Unable to retrieve the access token silently. error: {ex}"
                });
            }
        }
    }
}
