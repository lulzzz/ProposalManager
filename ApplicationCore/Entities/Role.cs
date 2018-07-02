// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace ApplicationCore.Entities
{
    public class Role : BaseEntity<Role>
    {
        /// <summary>
        /// Role display name
        /// </summary>
        [JsonProperty("displayName", Order = 2)]
        public string DisplayName { get; set; }

        /// <summary>
        /// AD Group Id display name
        /// </summary>
        [JsonProperty("adGroupName", Order = 3)]
        public string AdGroupName { get; set; }

        /// <summary>
        /// Represents the empty client. This field is read-only.
        /// </summary>
        public static Role Empty
        {
            get => new Role
            {
                Id = String.Empty,
                DisplayName = String.Empty,
                AdGroupName = String.Empty
            };
        }
    }
}
