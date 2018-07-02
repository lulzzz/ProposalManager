// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ApplicationCore.Artifacts;
using ApplicationCore.Helpers;
using ApplicationCore.Serialization;
using Newtonsoft.Json;
using WebReact.Models;
using WebReact.Serialization;
using WebReact.ViewModels;

namespace WebReact.Models
{
    public class OpportunityIndexModel
    {
        public OpportunityIndexModel()
        {
            Id = String.Empty;
            DisplayName = String.Empty;
            OpportunityState = OpportunityStateModel.NoneEmpty;
            Customer = new CustomerModel();
            DealSize = 0.0;
            OpenedDate = DateTimeOffset.MinValue;
        }

        /// <summary>
        /// Unique identifier of the artifact
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Unique identifier of the artifact
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }


        // Metadata
        [JsonConverter(typeof(OpportunityStateModelConverter))]
        [JsonProperty("opportunityState")]
        public OpportunityStateModel OpportunityState { get; set; }

        [JsonProperty("customer")]
        public CustomerModel Customer { get; set; }

        [JsonProperty("dealSize")]
        public double DealSize { get; set; }

        [JsonProperty("openedDate")]
        public DateTimeOffset OpenedDate { get; set; }
    }
}
