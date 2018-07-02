// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ApplicationCore;
using ApplicationCore.Helpers;
using Newtonsoft.Json;

namespace WebReact.Models
{
    public class DocumentModel
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("contentType")]
        public ContentType ContentType { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("reference")]
        public string Reference { get; set; }

        [JsonProperty("version")]
        public string Version { get; set; }

        // Metadata
        [JsonProperty("documentUri")]
        public string DocumentUri { get; set; }

        [JsonProperty("notes")]
        public IList<NoteModel> Notes { get; set; }

        [JsonProperty("category")]
        public CategoryModel Category { get; set; }

        [JsonProperty("tags")]
        public string Tags { get; set; }

        // Content
        [JsonProperty("content")]
        public DocumentContentModel Content { get; set; }
    }

    public class DocumentContentModel : ValueObject
    {
    }
}
