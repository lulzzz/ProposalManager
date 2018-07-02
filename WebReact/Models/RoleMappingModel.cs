// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace WebReact.Models
{
    public class RoleMappingModel
    {
        /// <summary>
        /// Role mapping identifier
        /// </summary>
        /// <value>Unique ID to identify the model data</value>
        [JsonProperty("id", Order = 1)]
        public string Id { get; set; }

        /// <summary>
        /// Role name
        /// </summary>
        [JsonProperty("roleName", Order = 2)]
        public string RoleName { get; set; }

        /// <summary>
        /// AD Group display name
        /// </summary>
        [JsonProperty("adGroupName", Order = 3)]
        public string AdGroupName { get; set; }

        /// <summary>
        /// AD Group Id 
        /// </summary>
        [JsonProperty("adGroupId", Order = 4)]
        public string AdGroupId { get; set; }

        /// <summary>
        /// Process Step 
        /// </summary>
        [JsonProperty("processStep", Order = 5)]
        public string ProcessStep { get; set; }

        /// <summary>
        /// Process Type 
        /// </summary>
        [JsonProperty("processType", Order = 6)]
        public string ProcessType { get; set; }

        /// <summary>
        /// Channel 
        /// </summary>
        [JsonProperty("channel", Order = 7)]
        public string Channel { get; set; }


        /// <summary>
        /// Represents the empty client. This field is read-only.
        /// </summary>
        public static RoleMappingModel Empty
        {
            get => new RoleMappingModel
            {
                Id = String.Empty,
                RoleName = String.Empty,
                AdGroupName = String.Empty,
                AdGroupId = String.Empty,
                ProcessStep = String.Empty,
                ProcessType = String.Empty,
                Channel = String.Empty
            };
        }
    }
}
