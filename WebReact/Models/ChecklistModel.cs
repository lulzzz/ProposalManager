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

namespace WebReact.Models
{
    public class ChecklistModel
    {
        public ChecklistModel()
        {
            Id = String.Empty;
            ChecklistChannel = String.Empty;
            ChecklistStatus = ActionStatus.NotStarted;
            ChecklistTaskList = new List<ChecklistTaskModel>();
        }

        [JsonProperty("id")]
        public string Id { get; set; }

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
        public IList<ChecklistTaskModel> ChecklistTaskList { get; set; }
    }

    public class ChecklistTaskModel
    {
        public ChecklistTaskModel()
        {
            Id = String.Empty;
            ChecklistItem = String.Empty;
            Completed = false;
            FileUri = String.Empty;
        }


        /// <summary>
        /// Checklist Id item
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Checklist task item
        /// </summary>
        [JsonProperty("checklistItem")]
        public string ChecklistItem { get; set; }

        /// <summary>
        /// Checklist task item completed flag
        /// </summary>
        [JsonProperty("completed")]
        public bool Completed { get; set; }

        /// <summary>
        /// Checklist task item
        /// </summary>
        [JsonProperty("fileUri")]
        public string FileUri { get; set; }
    }
}
