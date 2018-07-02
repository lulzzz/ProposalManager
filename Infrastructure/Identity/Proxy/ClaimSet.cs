// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System.Collections.Generic;

namespace Infrastructure.Identity.Proxy
{
    public class ClaimSet
    {
        public string ServiceName { get; set; }
        public Dictionary<string, string> Claims { get; set; }
        public ClaimSet InnerClaimSet { get; set; }
    }
}
