// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;

namespace ApplicationCore
{
    /// <summary>
    /// Settings relative to the AzureAD applications involved in this Web Application
    /// These are deserialized from the AzureAD section of the appsettings.json file
    /// </summary>
    public class AzureAdOptions
    {
        /// <summary>
        /// ClientId (Application Id) of this Web Application
        /// </summary>
        public string AppId { get; set; }

        /// <summary>
        /// ClientId (Application Id) of this Web Application
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// Client Secret (Application password) added in the Azure portal in the Keys section for the application
        /// </summary>
        public string ClientSecret { get; set; }

        /// <summary>
        /// Azure AD Cloud instance
        /// </summary>
        public string Instance { get; set; }

        /// <summary>
        ///  domain of your tenant, e.g. contoso.onmicrosoft.com
        /// </summary>
        public string Domain { get; set; }

        /// <summary>
        /// Tenant Id, as obtained from the Azure portal:
        /// (Select 'Endpoints' from the 'App registrations' blade and use the GUID in any of the URLs)
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// URL on which this Web App will be called back by Azure AD (normally "/signin-oidc")
        /// </summary>
        public string CallbackPath { get; set; }

        public string Authority => $"{Instance}{TenantId}";

        public string BaseUrl { get; set; }

        public string Scopes { get; set; }

        public string GraphResourceId { get; set; }

        public string GraphScopes { get; set; }
    }
}
