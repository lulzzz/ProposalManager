// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace CommLendingWeb.Helpers
{
	public class GraphSdkHelper : IGraphSdkHelper
    {
        private readonly IGraphAuthProvider authProvider;
        private GraphServiceClient graphClient;

        public GraphSdkHelper(IGraphAuthProvider authProvider)
        {
            this.authProvider = authProvider;
        }

        // Get an authenticated Microsoft Graph Service client.
        public GraphServiceClient GetAuthenticatedClient()
        {
            graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
                async requestMessage =>
                {
					// Passing tenant ID to the sample auth provider to use as a cache key
					var accessToken = await authProvider.GetTokenOnBehalfOfAsync();

					// Append the access token to the request
					requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                }));

            return graphClient;
        }

		public async Task<string> GetProposalManagerToken()
		{
			var token = await authProvider.GetProposalManagerTokenOnBehalfOfAsync();
			return token;
		}

		public async Task<string> GetGraphToken()
		{
			var token = await authProvider.GetTokenOnBehalfOfAsync();
			return token;
		}
    }
    public interface IGraphSdkHelper
    {
		GraphServiceClient GetAuthenticatedClient();
		Task<string> GetProposalManagerToken();
    }
}
