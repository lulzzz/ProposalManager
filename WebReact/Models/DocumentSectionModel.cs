// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ApplicationCore;
using ApplicationCore.Serialization;
using Newtonsoft.Json;
using WebReact.ViewModels;

namespace WebReact.Models
{
    public class DocumentSectionModel
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("owner")]
        public UserProfileViewModel Owner { get; set; }

        [JsonConverter(typeof(StatusConverter))]
        [JsonProperty("sectionStatus")]
        public ActionStatus SectionStatus { get; set; }

        [JsonProperty("lastModifiedDateTime")]
        public DateTimeOffset LastModifiedDateTime { get; set; }

        [JsonProperty("subSectionId")]
        public string SubSectionId { get; set; }

        [JsonProperty("assignedTo")]
        public UserProfileViewModel AssignedTo { get; set; }

        [JsonProperty("task")]
        public string Task { get; set; }
    }
}
