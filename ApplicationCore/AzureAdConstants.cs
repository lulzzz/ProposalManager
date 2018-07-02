// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;

namespace ApplicationCore
{
    public class AzureAdConstants
    {
        public static string TenantIdClaimType = "http://schemas.microsoft.com/identity/claims/tenantid";
        public static string ObjectIdClaimType = "http://schemas.microsoft.com/identity/claims/objectidentifier";
        public static string Common = "common";
        public static string AdminConsent = "admin_consent";
        public static string Issuer = "iss";
    }
}
