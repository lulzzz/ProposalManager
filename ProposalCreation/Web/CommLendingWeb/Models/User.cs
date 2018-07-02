// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using Newtonsoft.Json;

namespace CommLendingWeb.Models
{
	public class User : Entity
    {
		public string DisplayName { get; set; }
		public string UserPrincipalName { get; set; }
		[JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
		public string GivenName { get; set; }
		[JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
		public string Surname { get; set; }
		[JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
		public string Mail { get; set; }
	}
}
