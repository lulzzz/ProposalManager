// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using Newtonsoft.Json;
using System;

namespace CommLendingWeb.Models
{
	public class BaseItem : Entity
    {
		public DateTimeOffset? CreatedDateTime { get; set; }
		[JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
		public string Description { get; set; }
		public DateTimeOffset? LastModifiedDateTime { get; set; }
		[JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
		public string Name { get; set; }
		public string WebUrl { get; set; }
		public User CreatedByUser { get; set; }
		public User LastModifiedByUser { get; set; }
	}
}
