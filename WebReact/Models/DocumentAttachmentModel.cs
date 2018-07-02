// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace WebReact.Models
{
    public class DocumentAttachmentModel
    {
        public DocumentAttachmentModel()
        {
            Id = String.Empty;
            Note = String.Empty;
            Category = new CategoryModel();
            Tags = String.Empty;
            DocumentUri = String.Empty;
        }

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("fileName")]
        public string FileName { get; set; }

        /// <summary>
        /// Document notes
        /// </summary>
        [JsonProperty("note")]
        public string Note { get; set; }

        [JsonProperty("category")]
        public CategoryModel Category { get; set; }

        [JsonProperty("tags")]
        public string Tags { get; set; }

        /// <summary>
        /// Document location Uri
        /// </summary>
        [JsonProperty("documentUri")]
        public string DocumentUri { get; set; }
    }
}
