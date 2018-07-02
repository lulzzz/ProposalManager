// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebReact.Models
{
    public class RoleModel
    {
        public RoleModel()
        {
            Id = String.Empty;
            DisplayName = String.Empty;
            AdGroupName = String.Empty;
        }

        /// <summary>
        /// Role identifier
        /// </summary>
        /// <value>Unique ID to identify the model data</value>
        [JsonProperty("id", Order = 1)]
        public string Id { get; set; }

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
    }
}
