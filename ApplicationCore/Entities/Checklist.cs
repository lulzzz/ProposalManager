// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using ApplicationCore.Helpers;
using ApplicationCore.Serialization;
using Newtonsoft.Json;

namespace ApplicationCore.Entities
{
    public class Checklist : BaseEntity<Checklist>
    {
        /// <summary>
        /// Checklist overall status
        /// </summary>
        [JsonProperty("checklistChannel", Order = 2)]
        public string ChecklistChannel { get; set; }

        /// <summary>
        /// Checklist overall status
        /// </summary>
        [JsonConverter(typeof(StatusConverter))]
        [JsonProperty("checklistStatus", Order = 3)]
        public ActionStatus ChecklistStatus { get; set; }

        /// <summary>
        /// Checklist tasks list
        /// </summary>
        [JsonProperty("checklistTaskList", Order = 4)]
        public IList<ChecklistTask> ChecklistTaskList { get; set; }

        /// <summary>
        /// Represents the empty user profile. This field is read-only.
        /// </summary>
        public static Checklist Empty
        {
            get => new Checklist
            {
                Id = String.Empty,
                ChecklistStatus = ActionStatus.NotStarted,
                ChecklistTaskList = new List<ChecklistTask>()
            };
        }
    }

    public class ChecklistTask
    {
        public ChecklistTask()
        {
            Id = String.Empty;
            ChecklistItem = String.Empty;
            Completed = false;
            FileUri = String.Empty;
        }

        /// <summary>
        /// Checklist Id item
        /// </summary>
        [JsonProperty("id", Order = 1)]
        public string Id { get; set; }

        /// <summary>
        /// Checklist task item
        /// </summary>
        [JsonProperty("checklistItem", Order = 2)]
        public string ChecklistItem { get; set; }

        /// <summary>
        /// Checklist task item completed flag
        /// </summary>
        [JsonProperty("completed", Order = 3)]
        public bool Completed { get; set; }

        /// <summary>
        /// Checklist task item
        /// </summary>
        [JsonProperty("fileUri", Order = 4)]
        public string FileUri { get; set; }
    }
}
