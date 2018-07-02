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
    public class DocumentAttachment : BaseEntity<DocumentAttachment>
    {
        [JsonProperty("fileName")]
        public string FileName { get; set; }

        /// <summary>
        /// Document notes
        /// </summary>
        [JsonProperty("note")]
        public string Note { get; set; }

        [JsonProperty("category")]
        public Category Category { get; set; }

        [JsonProperty("tags")]
        public string Tags { get; set; }

        /// <summary>
        /// Document location Uri
        /// </summary>
        [JsonProperty("documentUri")]
        public string DocumentUri { get; set; }

        /// <summary>
        /// Represents the empty opportunity. This field is read-only.
        /// </summary>
        public static DocumentAttachment Empty
        {
            get => new DocumentAttachment
            {
                Id = String.Empty,
                Note = String.Empty,
                Category = Category.Empty,
                Tags = String.Empty,
                DocumentUri = String.Empty
            };
        }
    }
}
