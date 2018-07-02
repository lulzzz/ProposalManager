// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ApplicationCore.Artifacts;
using ApplicationCore.Entities;
using Newtonsoft.Json;

namespace WebReact.ViewModels
{
    public class DocumentViewModel
    {
        public DocumentViewModel()
        {
            // TODO: Init to empty and add empty static method
        }

        /// <summary>
        /// Unique identifier of the artifact
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Display name of the document
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("reference")]
        public string Reference { get; set; }

        [JsonProperty("version")]
        public string Version { get; set; }

        /// <summary>
        /// Document name
        /// </summary>
        [JsonProperty("documentName")]
        public String DocumentName { get; set; }

        /// <summary>
        /// Document location Uri (document name in UI)
        /// </summary>
        [JsonProperty("documentUri")]
        public String DocumentUri { get; set; }

        [JsonProperty("notes")]
        public String Notes { get; set; }

        [JsonProperty("category")]
        public Category Category { get; set; }

        [JsonProperty("tags")]
        public String Tags { get; set; }
    }
}
