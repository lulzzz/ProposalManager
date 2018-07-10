// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using Owin;
using System.Data.SqlClient;
using Microsoft.Azure;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SqlServer.Management.AlwaysEncrypted.AzureKeyVaultProvider;

namespace SmartLink.Web
{
    public partial class Startup
    {
        private static ClientCredential _clientCredential;

        public void Configuration(IAppBuilder app)
        {
            InitializeAzureKeyVaultProvider();
            ConfigureAuth(app);
        }

        static void InitializeAzureKeyVaultProvider()
        {
            string clientId = CloudConfigurationManager.GetSetting("ida:ClientId");
            string clientSecret = CloudConfigurationManager.GetSetting("ida:ClientSecret");
            _clientCredential = new ClientCredential(clientId, clientSecret);
            SqlColumnEncryptionAzureKeyVaultProvider azureKeyVaultProvider = new SqlColumnEncryptionAzureKeyVaultProvider(GetToken);
            Dictionary<string, SqlColumnEncryptionKeyStoreProvider> providers = new Dictionary<string, SqlColumnEncryptionKeyStoreProvider>();
            providers.Add(SqlColumnEncryptionAzureKeyVaultProvider.ProviderName, azureKeyVaultProvider);
            SqlConnection.RegisterColumnEncryptionKeyStoreProviders(providers);
        }

        /// <summary>
        /// A callback to redeem the application id/key for a token for Azure Key Vault
        /// </summary>
        public async static Task<string> GetToken(string authority, string resource, string scope)
        {
            var authContext = new AuthenticationContext(authority);
            AuthenticationResult result = await authContext.AcquireTokenAsync(resource, _clientCredential);

            if (result == null)
                throw new InvalidOperationException("Failed to obtain the JWT token");

            return result.AccessToken;
        }
    }
}