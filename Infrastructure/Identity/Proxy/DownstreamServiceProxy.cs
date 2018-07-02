// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using ApplicationCore;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading.Tasks;

namespace Infrastructure.Identity.Proxy
{
    public class DownstreamServiceProxy
    {
        private readonly AzureAdOptions authOptions;
        private readonly DownstreamServiceProxyOptions serviceOptions;
        private readonly IHttpContextAccessor httpContextAccessor;

        public DownstreamServiceProxy(
            IOptions<AzureAdOptions> authOptions, 
            IOptions<DownstreamServiceProxyOptions> serviceOptions,
            IHttpContextAccessor httpContextAccessor)
        {
            this.authOptions = authOptions.Value;
            this.serviceOptions = serviceOptions.Value;
            this.httpContextAccessor = httpContextAccessor;
        }

        public async Task<ClaimSet> GetClaimSetAsync()
        {
            var client = new HttpClient { BaseAddress = new Uri(serviceOptions.BaseUrl, UriKind.Absolute) };
            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", await GetAccessTokenAsync());

            var payload = await client.GetStringAsync("api/claims");
            return JsonConvert.DeserializeObject<ClaimSet>(payload);
        }

        private async Task<string> GetAccessTokenAsync()
        {
            var credential = new ClientCredential(authOptions.ClientId, authOptions.ClientSecret);
            var authenticationContext = new AuthenticationContext(authOptions.Authority);

            var originalToken = await httpContextAccessor.HttpContext.GetTokenAsync("access_token");
            var userName = httpContextAccessor.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value ??
                httpContextAccessor.HttpContext.User.FindFirst(ClaimTypes.Name)?.Value;

            var userAssertion = new UserAssertion(originalToken, 
                "urn:ietf:params:oauth:grant-type:jwt-bearer", userName);

            var result = await authenticationContext.AcquireTokenAsync(serviceOptions.Resource,
                credential, userAssertion);

            return result.AccessToken;
        }
    }
}
