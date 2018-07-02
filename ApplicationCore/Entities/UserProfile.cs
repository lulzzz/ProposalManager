// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using ApplicationCore.Serialization;
using Newtonsoft.Json;

namespace ApplicationCore.Entities
{
    public class UserProfile : BaseEntity<UserProfile>
    {
        /// <summary>
        /// User display name
        /// </summary>
        [JsonProperty("displayName", Order = 1)]
        public string DisplayName { get; set; }

        /// <summary>
        /// The values for the user profile fields
        /// </summary>
        [JsonProperty("fields", Order = 2)]
        public UserProfileFields Fields { get; set; }

        /// <summary>
        /// Represents the empty user profile. This field is read-only.
        /// </summary>
        public static UserProfile Empty {
            get => new UserProfile {
                Id = String.Empty,
                DisplayName = String.Empty,
                Fields = UserProfileFields.Empty
            };
        }
    }

    public class UserProfileFields
    {
        /// <summary>
        /// User email
        /// </summary>
        [JsonProperty("mail")]
        public string Mail { get; set; }

        /// <summary>
        /// User Principal Name
        /// </summary>
        [JsonProperty("userPrincipalName")]
        public string UserPrincipalName { get; set; }

        /// <summary>
        /// User title
        /// </summary>
        [JsonProperty("title")]
        public string Title { get; set; }

        /// <summary>
        /// User role
        /// </summary>
        //[JsonConverter(typeof(SmartEnumConverter))]
        [JsonProperty("userRoles")]
        public List<Role> UserRoles { get; set; }

        /// <summary>
        /// Represents the empty user profile. This field is read-only.
        /// </summary>
        public static UserProfileFields Empty
        {
            get => new UserProfileFields {
                Mail = String.Empty,
                UserPrincipalName = String.Empty,
                Title = String.Empty,
                UserRoles = new List<Role>()
            };
        }
    }
}
