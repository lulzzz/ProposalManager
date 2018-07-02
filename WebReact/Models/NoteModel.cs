// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebReact.ViewModels;

namespace WebReact.Models
{
    public class NoteModel
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Body of the note
        /// </summary>
        [JsonProperty("noteBody")]
        public string NoteBody { get; set; }

        /// <summary>
        /// Created date
        /// </summary>
        [JsonProperty("createdDateTime")]
        public DateTimeOffset CreatedDateTime { get; set; }

        /// <summary>
        /// User that created the note
        /// </summary>
        [JsonProperty("createdBy")]
        public UserProfileViewModel CreatedBy { get; set; }
    }
}
