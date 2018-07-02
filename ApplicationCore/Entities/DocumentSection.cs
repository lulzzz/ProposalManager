// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using ApplicationCore.Serialization;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace ApplicationCore.Entities
{
    public class DocumentSection : BaseEntity<DocumentSection>
    {
        /// <summary>
        /// Display name for the section
        /// </summary>
        [JsonProperty("displayName", Order = 2)]
        public string DisplayName { get; set; }

        /// <summary>
        /// Section owner
        /// </summary>
        [JsonProperty("owner", Order = 3)]
        public UserProfile Owner { get; set; }


        /// <summary>
        /// Section status
        /// </summary>
        [JsonConverter(typeof(StatusConverter))]
        [JsonProperty("sectionStatus", Order = 4)]
        public ActionStatus SectionStatus { get; set; }

        /// <summary>
        /// Last updated
        /// </summary>
        [JsonProperty("lastModifiedDateTime", Order = 5)]
        public DateTimeOffset LastModifiedDateTime { get; set; }

        /// <summary>
        /// Parent of the proposal Section. Empty means root
        /// </summary>
        [JsonProperty("subSectionId", Order = 6)]
        public string SubSectionId { get; set; }

        /// <summary>
        /// User which this section is assigned to.
        /// </summary>
        [JsonProperty("assignedTo", Order = 7)]
        public UserProfile AssignedTo { get; set; }

        /// <summary>
        /// Task text.
        /// </summary>
        [JsonProperty("task", Order = 8)]
        public string Task { get; set; }

        /// <summary>
        /// Represents the empty opportunity. This field is read-only.
        /// </summary>
        public static DocumentSection Empty
        {
            get => new DocumentSection
            {
                Id = String.Empty,
                DisplayName = String.Empty,
                Owner = UserProfile.Empty,
                SectionStatus = ActionStatus.NotStarted,
                LastModifiedDateTime = DateTimeOffset.MinValue,
                SubSectionId = String.Empty,
                AssignedTo = UserProfile.Empty,
                Task = String.Empty
            };
        }
    }
}
