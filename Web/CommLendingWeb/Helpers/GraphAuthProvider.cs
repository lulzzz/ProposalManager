// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using CommLendingWeb.Extensions;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using System;
using System.Security.Claims;
using System.Threading.Tasks;

namespace CommLendingWeb.Helpers
{
	public class GraphAuthProvider : IGraphAuthProvider
	{
		private readonly string appId;
		private readonly ClientCredential credential;
		private readonly string[] scopes;
		private readonly IHttpContextAccessor contextAccessor;
		private readonly string secret;
		private readonly string proposalManagerApiId;

		public GraphAuthProvider(IMemoryCache memoryCache, IConfiguration configuration, IHttpContextAccessor contextAccessor)
		{
			
			var azureOptions = new AzureAdOptions();
			configuration.Bind("AzureAd", azureOptions);

			appId = azureOptions.ClientId;
			credential = new ClientCredential(azureOptions.ClientSecret);
			scopes = azureOptions.GraphScopes.Split(new[] { ' ' });
			secret = azureOptions.ClientSecret;
			proposalManagerApiId = azureOptions.ProposalManagerApiId;
			this.contextAccessor = contextAccessor;
		}

		public async Task<string> GetProposalManagerTokenOnBehalfOfAsync()
		{
			try
			{
				const string siteUrl = "proposalcreation.azurewebsites.net";
				// Get the raw token that the add-in page received from the Office host.
				var bootstrapContext = ((ClaimsIdentity)contextAccessor.HttpContext.User.Identity).BootstrapContext.ToString();
				
				UserAssertion userAssertion = new UserAssertion(bootstrapContext, "urn:ietf:params:oauth:grant-type:jwt-bearer");

				// Get the access token for MS Graph. 
				ClientCredential clientCred = new ClientCredential(secret);
				ConfidentialClientApplication cca =
					new ConfidentialClientApplication(appId,
														$"https://{siteUrl}", clientCred, null, null);

				// The AcquireTokenOnBehalfOfAsync method will first look in the MSAL in memory cache for a
				// matching access token. Only if there isn't one, does it initiate the "on behalf of" flow
				// with the Azure AD V2 endpoint.
				var apiScope = $"api://{proposalManagerApiId}/access_as_user";
				AuthenticationResult result = await cca.AcquireTokenOnBehalfOfAsync(new[] { apiScope }, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
				return result.AccessToken;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public async Task<string> GetTokenOnBehalfOfAsync()
		{
			try
			{
				const string siteUrl = "proposalcreation.azurewebsites.net"; 
				// Get the raw token that the add-in page received from the Office host.
				var bootstrapContext = ((ClaimsIdentity)contextAccessor.HttpContext.User.Identity).BootstrapContext.ToString();

				UserAssertion userAssertion = new UserAssertion(bootstrapContext);

				// Get the access token for MS Graph. 
				ClientCredential clientCred = new ClientCredential(this.secret);
				ConfidentialClientApplication cca =
					new ConfidentialClientApplication(appId,
														$"https://{siteUrl}", clientCred, null, null);

				// The AcquireTokenOnBehalfOfAsync method will first look in the MSAL in memory cache for a
				// matching access token. Only if there isn't one, does it initiate the "on behalf of" flow
				// with the Azure AD V2 endpoint.
				AuthenticationResult result = await cca.AcquireTokenOnBehalfOfAsync(scopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
				return result.AccessToken;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}

	public interface IGraphAuthProvider
	{
		Task<string> GetTokenOnBehalfOfAsync();
		Task<string> GetProposalManagerTokenOnBehalfOfAsync();
	}
}