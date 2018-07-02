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
    public class CustomerDecision : BaseEntity<Category>
    {
        [JsonProperty("approved", Order = 2)]
        public bool Approved { get; set; }

        [JsonProperty("approvedDate", Order = 3)]
        public DateTimeOffset ApprovedDate { get; set; }

        [JsonProperty("loanDisbursed", Order = 4)]
        public DateTimeOffset LoanDisbursed { get; set; }

        public static CustomerDecision Empty
        {
            get => new CustomerDecision
            {
                Id = String.Empty,
                Approved = false,
                ApprovedDate = DateTimeOffset.MinValue,
                LoanDisbursed = DateTimeOffset.MinValue
            };
        }
    }
}
