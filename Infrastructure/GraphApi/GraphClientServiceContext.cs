// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using System.Security.Claims;
using System.Net.Http.Headers;
using Microsoft.Graph;
using ApplicationCore;
using ApplicationCore.Interfaces;
using ApplicationCore.Helpers;

namespace Infrastructure.GraphApi
{
    public class GraphClientUserContext : IGraphClientUserContext
    {
        private readonly IGraphAuthProvider _authProvider;
        private readonly IUserContext _userContext;
        private readonly GraphServiceClient _graphServiceClient;

        public GraphClientUserContext(
            IGraphAuthProvider authProvider,
            IUserContext userContext)
        {
            Guard.Against.Null(authProvider, nameof(authProvider));
            Guard.Against.Null(userContext, nameof(userContext));
            _authProvider = authProvider;
            _userContext = userContext;

            // Initialize the graph client given the chosen context
            if (_graphServiceClient == null)
            {
                _graphServiceClient = new GraphServiceClient(new DelegateAuthenticationProvider(
                async requestMessage =>
                {
                    // Passing tenant ID to the auth provider to use as a cache key
                    var accessToken = await _authProvider.GetUserAccessTokenAsync(User.FindFirst(AzureAdConstants.ObjectIdClaimType)?.Value);

                    // Append the access token to the request
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                }));
            }
        }

        /// <summary>
        /// Gets the <see cref="ClaimsPrincipal"/> for user associated with the executing action.
        /// </summary>
        public ClaimsPrincipal User => _userContext?.User;

        /// <summary>
        /// Graph Service client using the user context
        /// </summary>
        public GraphServiceClient GraphClient => _graphServiceClient;
    }

    public class GraphClientAppContext : IGraphClientAppContext
    {
        private readonly IGraphAuthProvider _authProvider;
        private readonly IUserContext _userContext;
        private readonly GraphServiceClient _graphServiceClient;

        public GraphClientAppContext(
            IGraphAuthProvider authProvider,
            IUserContext userContext)
        {
            _authProvider = authProvider ?? throw new ArgumentNullException(nameof(authProvider));
            _userContext = userContext ?? throw new ArgumentNullException(nameof(userContext));

            // Initialize the graph client given the chosen context
            if (_graphServiceClient == null)
            {
                _graphServiceClient = new GraphServiceClient(new DelegateAuthenticationProvider(
                async requestMessage =>
                {
                    // Passing tenant ID to the auth provider to use as a cache key
                    var accessToken = await _authProvider.GetAppAccessTokenAsync();

                    // Append the access token to the request
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                }));
            }
        }

        /// <summary>
        /// Gets the <see cref="ClaimsPrincipal"/> for user associated with the executing action.
        /// </summary>
        public ClaimsPrincipal User => _userContext?.User;

        /// <summary>
        /// Graph Service client using the application context
        /// </summary>
        public GraphServiceClient GraphClient => _graphServiceClient;
    }
}
