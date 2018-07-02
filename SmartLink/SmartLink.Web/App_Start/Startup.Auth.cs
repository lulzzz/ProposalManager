// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using Microsoft.Azure;
using Microsoft.Owin.Security.ActiveDirectory;
using Owin;
using System.IdentityModel.Tokens;

namespace SmartLink.Web
{
	public partial class Startup
    {
        private static string clientId = CloudConfigurationManager.GetSetting("ida:ClientId");
        private static string tenantId = CloudConfigurationManager.GetSetting("ida:TenantId");

        public void ConfigureAuth(IAppBuilder app)
        {
			app.UseWindowsAzureActiveDirectoryBearerAuthentication(new WindowsAzureActiveDirectoryBearerAuthenticationOptions
			{
				Tenant = tenantId,
				TokenValidationParameters = new TokenValidationParameters { SaveSigninToken = true, ValidAudience = clientId }
			});
		}
    }
}