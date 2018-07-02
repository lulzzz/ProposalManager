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
    public class Region : BaseEntity<Region>
    {
        /// <summary>
        /// Region display name
        /// </summary>
        [JsonProperty("name", Order = 2)]
        public string Name { get; set; }

        /// <summary>
        /// Represents the empty object. This field is read-only.
        /// </summary>
        public static Region Empty
        {
            get => new Region
            {
                Id = String.Empty,
                Name = String.Empty
            };
        }
    }
}
