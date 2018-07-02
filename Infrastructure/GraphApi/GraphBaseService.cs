// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;

using System.IO;
using System.Threading.Tasks;
using System.Security.Claims;
using System.Net.Http;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http.Headers;
using ApplicationCore;
using ApplicationCore.Interfaces;
using Infrastructure.Identity.Extensions;
using Microsoft.Extensions.Options;
using Microsoft.Extensions.Logging;
using Infrastructure.Services;
using ApplicationCore.Helpers;


// TODO: To be deprecated (delete this file) after initial test of new IoC in GraphServices (refactor of context)

namespace Infrastructure.GraphApi
{
    public abstract class GraphBaseService<T> : BaseService<T>
    {
        protected readonly IGraphAuthProvider _authProvider;
        protected readonly IUserContext _userContext;
        protected readonly IGraphClientContext _graphClientContext;
        protected GraphClientContext _currentClientContext;
        protected GraphServiceClient _graphServiceClient;
        protected GraphServiceClient _graphServiceAppClient;
        protected GraphServiceClient _graphServiceOnBehalfClient;


        public GraphBaseService(
            ILogger<T> logger,
            IOptions<AppOptions> appOptions,
            IGraphAuthProvider authProvider, 
            IUserContext userContext,
            IGraphClientContext graphClientContext) : base(logger, appOptions)
        {
            Guard.Against.Null(authProvider, nameof(authProvider));
            Guard.Against.Null(userContext, nameof(userContext));
            Guard.Against.Null(graphClientContext, nameof(graphClientContext));
            _authProvider = authProvider;
            _userContext = userContext;

            // Set current context to User by default (always default to least privilege)
            _currentClientContext = GraphClientContext.User;
        }


        /// <summary>
        /// Gets the <see cref="ClaimsPrincipal"/> for user associated with the executing action.
        /// </summary>
        public ClaimsPrincipal User => _userContext?.User;

        /// <summary>
        /// Graph Service client
        /// </summary>
        public GraphServiceClient GraphClient => _graphClientContext?.GraphClient;



        public GraphClientContext CurrentClientContext
        {
            get
            {
                return _currentClientContext;
            }

            set
            {
                _currentClientContext = value;
            }
        }

        /// <summary>
        /// Get the Graph Service client using the current user context
        /// </summary>
        /// <returns>Graph Service client</returns>
        //public GraphServiceClient GraphClient
        //{
        //    get
        //    {
        //        switch (_currentClientContext.Name)
        //        {
        //            case "Application":
        //                return GetAppServiceClient();
        //            case "OnBehalf":
        //                return GetOnBehalfServiceClient();
        //        }

        //        // By default return current user context
        //        return GetAuthenticatedClient();
        //    }
        //}

        /// <summary>
        /// Get an authenticated Microsoft Graph Service client using the current user context.
        /// </summary>
        /// <returns>Graph Service client</returns>
        public GraphServiceClient GetAuthenticatedClient()
        {
            return GetAuthenticatedClient(GraphClientContext.User, User.FindFirst(AzureAdConstants.ObjectIdClaimType)?.Value);
        }

        /// <summary>
        /// Get an authenticated Microsoft Graph Service client using the context of the specified user.
        /// </summary>
        /// <param name="userId">User identifier to get the token</param>
        /// <returns>Graph Service client</returns>
        public GraphServiceClient GetAuthenticatedClient(string userId)
        {
            return GetAuthenticatedClient(GraphClientContext.User, userId);
        }

        /// <summary>
        /// Get an authenticated Microsoft Graph Service client using the specified context
        /// </summary>
        /// <param name="graphClientContext">Defines the context for the GraphClient: User, Application or OnBehalf</param>
        /// <param name="userId">Optional: User identifier to get the token</param>
        /// <returns></returns>
        public GraphServiceClient GetAuthenticatedClient(GraphClientContext graphClientContext, string userId = "")
        {
            switch (graphClientContext.Name)
            {
                case "Application":
                    return GetAppServiceClient();
                case "OnBehalf":
                    return GetOnBehalfServiceClient();
            }

            // By default return current user context
            return GetUserServiceClient(userId);
        }


        // Private methods
        private GraphServiceClient GetUserServiceClient(string userId)
        {
            if (_graphServiceClient != null) return _graphServiceClient;
            _graphServiceClient = new GraphServiceClient(new DelegateAuthenticationProvider(
                async requestMessage =>
                {
                    // Passing tenant ID to the auth provider to use as a cache key
                    var accessToken = await _authProvider.GetUserAccessTokenAsync(userId);

                    // Append the access token to the request
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                }));

            return _graphServiceClient;
        }

        private GraphServiceClient GetAppServiceClient()
        {
            if (_graphServiceAppClient != null) return _graphServiceAppClient;
            _graphServiceAppClient = new GraphServiceClient(new DelegateAuthenticationProvider(
                async requestMessage =>
                {
                    // Passing tenant ID to the auth provider to use as a cache key
                    var accessToken = await _authProvider.GetAppAccessTokenAsync();

                    // Append the access token to the request
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                }));

            return _graphServiceAppClient;
        }

        private GraphServiceClient GetOnBehalfServiceClient()
        {
            // TODO
            return null;
        }
    }
}
