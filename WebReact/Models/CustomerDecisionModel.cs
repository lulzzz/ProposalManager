// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace WebReact.Models
{
    public class CustomerDecisionModel
    {
        [JsonProperty("id", Order = 1)]
        public string Id { get; set; }

        [JsonProperty("approved", Order = 2)]
        public bool Approved { get; set; }

        [JsonProperty("approvedDate", Order = 3)]
        public DateTimeOffset ApprovedDate { get; set; }

        [JsonProperty("loanDisbursed", Order = 4)]
        public DateTimeOffset LoanDisbursed { get; set; }
    }
}
