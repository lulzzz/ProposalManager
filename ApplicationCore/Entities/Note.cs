// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace ApplicationCore.Entities
{
    public class Note : BaseEntity<Note>
    {
        /// <summary>
        /// Body of the note
        /// </summary>
        [JsonProperty("noteBody", Order = 2)]
        public string NoteBody { get; set; }

        /// <summary>
        /// Created date
        /// </summary>
        [JsonProperty("createdDateTime", Order = 3)]
        public DateTimeOffset CreatedDateTime { get; set; }

        /// <summary>
        /// User that created the note
        /// </summary>
        [JsonProperty("createdBy", Order = 4)]
        public UserProfile CreatedBy { get; set; }

        /// <summary>
        /// Represents the empty opportunity. This field is read-only.
        /// </summary>
        public static Note Empty
        {
            get => new Note
            {
                Id = String.Empty,
                NoteBody = String.Empty,
                CreatedDateTime = DateTimeOffset.MinValue,
                CreatedBy = UserProfile.Empty
            };
        } 
    }
}
