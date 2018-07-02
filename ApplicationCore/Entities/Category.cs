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
    public class Category : BaseEntity<Category>
    {
        /// <summary>
        /// Category display name
        /// </summary>
        [JsonProperty("name", Order = 2)]
        public string Name { get; set; }

        /// <summary>
        /// Represents the empty object. This field is read-only.
        /// </summary>
        public static Category Empty
        {
            get => new Category
            {
                Id = String.Empty,
                Name = String.Empty
            };
        }
    }
}
