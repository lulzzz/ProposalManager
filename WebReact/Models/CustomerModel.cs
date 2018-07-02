// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace WebReact.Models
{
    public class CustomerModel
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Customer display name
        /// </summary>
        [JsonProperty("displayName", Order = 2)]
        public string DisplayName { get; set; }

        /// <summary>
        /// Reference ID of a customer to associate in a back en dsystem containing customer entity
        /// </summary>
        [JsonProperty("referenceId", Order = 3)]
        public string ReferenceId { get; set; }
    }
}
