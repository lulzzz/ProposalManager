// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public interface IConfigService
    {
        string ClientId { get;}
        string ClientSecret { get; }
        string WebJobClientId { get; }
        string AzureAdInstance { get; }
        string AzureAdTenantId { get; }
        string GraphResourceUrl{ get; }
        string AzureAdGraphResourceURL { get; }
        string AzureAdAuthority { get; }
        string ClaimTypeObjectIdentifier { get; }
        string AzureWebJobsStorage { get; }
        string AzureWebJobDashboard { get; }
        string SharePointUrl { get; }
        string CertificateFile { get; }
        string CertificatePassword { get; }
        string SendGridMessageUserName { get; }
        string SendGridMessagePassword { get; }
        string SendGridMessageFromAddress { get; }
        string SendGridMessageFromDisplayName { get; }
        string[] SendGridMessageToAddress { get; }
        string DatabaseConnectionString { get; }
   }
}
