// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using ApplicationCore;
using ApplicationCore.Serialization;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebReact.ViewModels;

namespace WebReact.Models
{
    public class TeamMemberModel
    {
        public TeamMemberModel()
        {
            Id = String.Empty;
            DisplayName = String.Empty;
            Mail = String.Empty;
            UserPrincipalName = String.Empty;
            Title = String.Empty;
            Status = ActionStatus.NotStarted;
            AssignedRole = new RoleModel();
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
        /// Status in the context of an instance of an opportunity
        /// </summary>
        [JsonConverter(typeof(StatusConverter))]
        [JsonProperty("status")]
        public ActionStatus Status { get; set; }

        [JsonProperty("assignedRole")]
        public RoleModel AssignedRole { get; set; }
    }
}
