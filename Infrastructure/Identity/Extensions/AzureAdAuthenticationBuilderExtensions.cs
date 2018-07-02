// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authentication;

using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Identity.Client;
using ApplicationCore;
using Microsoft.AspNetCore.Authentication.JwtBearer;

namespace Infrastructure.Identity.Extensions
{
    /// <summary>
    /// Extensions for the AzureAdAuthenticationBuilder
    /// </summary>
    public static class AzureAdAuthenticationBuilderExtensions
    {
        // Extensions for authenticating using barer token (user already authenticated in client) NOTE: To use this the corresponding initialization is needed in startup
        #region Barer token (for incoming webapi calls)
        public static AuthenticationBuilder AddAzureAdBearer(this AuthenticationBuilder builder)
            => builder.AddAzureAdBearer(_ => { });

        public static AuthenticationBuilder AddAzureAdBearer(this AuthenticationBuilder builder, Action<AzureAdOptions> configureOptions)
        {
            builder.Services.Configure(configureOptions);
            builder.Services.AddSingleton<IConfigureOptions<JwtBearerOptions>, ConfigureAzureAdBearerOptions>();
            builder.AddJwtBearer();
            return builder;
        }

        private class ConfigureAzureAdBearerOptions : IConfigureNamedOptions<JwtBearerOptions>
        {
            private readonly AzureAdOptions _azureOptions;

            public ConfigureAzureAdBearerOptions(IOptions<AzureAdOptions> azureOptions)
            {
                _azureOptions = azureOptions.Value;
            }

            public void Configure(string name, JwtBearerOptions options)
            {
                options.Audience = _azureOptions.ClientId;
                options.Authority = $"{_azureOptions.Instance}{_azureOptions.TenantId}";

                options.TokenValidationParameters = new TokenValidationParameters
                {
                    // Instead of using the default validation (validating against a single issuer value, as we do in line of business apps),
                    // we inject our own multitenant validation logic
                    ValidateIssuer = false,

                    // If the app is meant to be accessed by entire organizations, add your issuer validation logic here.
                    //IssuerValidator = (issuer, securityToken, validationParameters) =>
                    //{
                    //    if (myIssuerValidationLogic(issuer)) return issuer;
                    //    return String.Empty;
                    //}
                };
                options.Events = new JwtBearerEvents
                {
                    OnTokenValidated = TokenValidated,
                    OnAuthenticationFailed = AuthenticationFailed,
                    //OnMessageReceived = MessageReceived
                };
                options.SaveToken = true;

                options.Validate();
            }

            public void Configure(JwtBearerOptions options)
            {
                Configure(Options.DefaultName, options);
            }

            // MessageReceived event
            private Task MessageReceived(Microsoft.AspNetCore.Authentication.JwtBearer.MessageReceivedContext context)
            {
                // If no token found, no further work possible
                if (string.IsNullOrEmpty(context.Token))
                {
                    //return AuthenticateResult.NoResult();
                }

                return Task.FromResult(0);
            }

            // TokenValidated event
            private Task TokenValidated(Microsoft.AspNetCore.Authentication.JwtBearer.TokenValidatedContext context)
            {
                /* ---------------------   
                // Replace this with your logic to validate the issuer/tenant
                   ---------------------       
                // Retriever caller data from the incoming principal
                string issuer = context.SecurityToken.Issuer;
                string subject = context.SecurityToken.Subject;
                string tenantID = context.Ticket.Principal.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;

                // Build a dictionary of approved tenants
                IEnumerable<string> approvedTenantIds = new List<string>
                {
                    "<Your tenantID>",
                    "9188040d-6c67-4c5b-b112-36a304b66dad" // MSA Tenant
                };

                if (!approvedTenantIds.Contains(tenantID))
                    throw new SecurityTokenValidationException();
                  --------------------- */

                // Store the token in the token cache

                return Task.FromResult(0);
            }

            // Handle sign-in errors differently than generic errors.
            private Task AuthenticationFailed(Microsoft.AspNetCore.Authentication.JwtBearer.AuthenticationFailedContext context)
            {
                //context.HandleResponse();
                //context.Response.Redirect("/Home/Error?message=" + context.Failure.Message);
                //context.Response.Redirect("/Home/Error?message=");
                return Task.FromResult(0);
            }
        }
        #endregion



        // Extenions for autheticating the user in server side NOTE: To use this the corresponding initialization is needed in startup
        #region Cookie token (for server side based client)
        public static AuthenticationBuilder AddAzureAd(this AuthenticationBuilder builder)
            => builder.AddAzureAd(_ => { });

        public static AuthenticationBuilder AddAzureAd(this AuthenticationBuilder builder, Action<AzureAdOptions> configureOptions)
        {
            builder.Services.Configure(configureOptions);
            builder.Services.AddSingleton<IConfigureOptions<OpenIdConnectOptions>, ConfigureAzureAdOptions>();
            builder.AddOpenIdConnect();
            return builder;
        }

        public class ConfigureAzureAdOptions : IConfigureNamedOptions<OpenIdConnectOptions>
        {
            private readonly AzureAdOptions _azureOptions;

            public AzureAdOptions GetAzureAdOptions() => _azureOptions;

            public ConfigureAzureAdOptions(IOptions<AzureAdOptions> azureOptions)
            {
                _azureOptions = azureOptions.Value;
            }

            public void Configure(string name, OpenIdConnectOptions options)
            {
                options.ClientId = _azureOptions.ClientId;
                options.Authority = $"{_azureOptions.Instance}common/v2.0";
                options.UseTokenLifetime = true;
                options.CallbackPath = _azureOptions.CallbackPath;
                options.RequireHttpsMetadata = false;
                options.ResponseType = OpenIdConnectResponseType.CodeIdToken;
                var allScopes = $"{_azureOptions.Scopes} {_azureOptions.GraphScopes}".Split(new[] { ' ' });
                foreach (var scope in allScopes) { options.Scope.Add(scope); }

                options.TokenValidationParameters = new TokenValidationParameters
                {
                    // Instead of using the default validation (validating against a single issuer value, as we do in line of business apps),
                    // we inject our own multitenant validation logic
                    ValidateIssuer = false,

                    // If the app is meant to be accessed by entire organizations, add your issuer validation logic here.
                    //IssuerValidator = (issuer, securityToken, validationParameters) => {
                    //    if (myIssuerValidationLogic(issuer)) return issuer;
                    //}
                };

                options.Events = new OpenIdConnectEvents
                {
                    OnTicketReceived = context =>
                    {
                        // If your authentication logic is based on users then add your logic here
                        return Task.CompletedTask;
                    },
                    OnAuthenticationFailed = context =>
                    {
                        context.Response.Redirect("/Home/Error");
                        context.HandleResponse(); // Suppress the exception
                        return Task.CompletedTask;
                    },
                    OnAuthorizationCodeReceived = async (context) =>
                    {
                        var code = context.ProtocolMessage.Code;
                        var identifier = context.Principal.FindFirst(AzureAdConstants.ObjectIdClaimType).Value;
                        var memoryCache = context.HttpContext.RequestServices.GetRequiredService<IMemoryCache>();
                        var graphScopes = _azureOptions.GraphScopes.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                        var cca = new ConfidentialClientApplication(
                            _azureOptions.ClientId,
                            _azureOptions.BaseUrl + _azureOptions.CallbackPath,
                            new ClientCredential(_azureOptions.ClientSecret),
                            new SessionTokenCache(identifier, memoryCache).GetCacheInstance(),
                            null);
                        var result = await cca.AcquireTokenByAuthorizationCodeAsync(code, graphScopes);

                        // Check whether the login is from the MSA tenant. 
                        // The sample uses this attribute to disable UI buttons for unsupported operations when the user is logged in with an MSA account.
                        var currentTenantId = context.Principal.FindFirst(AzureAdConstants.TenantIdClaimType).Value;
                        if (currentTenantId == "9188040d-6c67-4c5b-b112-36a304b66dad")
                        {
                            // MSA (Microsoft Account) is used to log in
                        }

                        context.HandleCodeRedemption(result.AccessToken, result.IdToken);
                    },
                    // If your application needs to do authenticate single users, add your user validation below.
                    //OnTokenValidated = context =>
                    //{
                    //    return myUserValidationLogic(context.Ticket.Principal);
                    //}
                };
            }

            public void Configure(OpenIdConnectOptions options)
            {
                Configure(Options.DefaultName, options);
            }
        }
        #endregion
    }
}
