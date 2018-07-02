// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using ApplicationCore.Artifacts;
using ApplicationCore.Entities;
using Newtonsoft.Json.Converters;
using Infrastructure.Serialization;
using ApplicationCore.Serialization;
using WebReact.Models;

namespace WebReact.ViewModels
{
    public class UserProfileViewModel
    {
        public UserProfileViewModel()
        {
            Id = String.Empty;
            DisplayName = String.Empty;
            Mail = String.Empty;
            UserPrincipalName = String.Empty;
            Title = String.Empty;
            UserRoles = new List<RoleModel>();
        }

        /// <summary>
        /// Unique identifier of the artifact
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// User display name
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

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
        public List<RoleModel> UserRoles { get; set; }
    }
}
