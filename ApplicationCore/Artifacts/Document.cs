// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using ApplicationCore.Entities;
using ApplicationCore.Helpers;
using Newtonsoft.Json;

namespace ApplicationCore.Artifacts
{
    public class Document : BaseArtifact<Document>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public Document()
        {
            ContentType = ContentType.Document;
            Version = "1.0";
        }

        /// <summary>
        /// Content type of the document
        /// </summary>
        [JsonProperty("contentType")]
        public new ContentType ContentType { get; set; }

        /// <summary>
        /// Metadata of the document
        /// </summary>
        [JsonProperty("metadata")]
        public new DocumentMetadata Metadata { get; set; }

        /// <summary>
        /// Content of the document
        /// </summary>
        [JsonProperty("content")]
        public new DocumentContent Content { get; set; }

        /// <summary>
        /// Represents the empty document. This field is read-only.
        /// </summary>
        public static Document Empty
        {
            get => new Document
            {
                Id = String.Empty,
                DisplayName = String.Empty,
                Reference = String.Empty,
                ContentType = ContentType.Document,
                Version = "1.0",
                Metadata = DocumentMetadata.Empty,
                Content = DocumentContent.Empty
            };
        }
    }

    public class DocumentMetadata
    {
        /// <summary>
        /// Document location Uri
        /// </summary>
        [JsonProperty("documentUri")]
        public string DocumentUri { get; set; }

        /// <summary>
        /// Document notes
        /// </summary>
        [JsonProperty("notes")]
        public IList<Note> Notes { get; set; }

        [JsonProperty("category")]
        public Category Category { get; set; }

        [JsonProperty("tags")]
        public string Tags { get; set; }

        /// <summary>
        /// Represents the empty document. This field is read-only.
        /// </summary>
        public static DocumentMetadata Empty
        {
            get => new DocumentMetadata
            {
                DocumentUri = String.Empty,
                Notes = new List<Note>(),
                Category = Category.Empty,
                Tags = String.Empty
            };
        }
    }

    public class DocumentContent : ValueObject
    {
        /// <summary>
        /// Represents the empty document. This field is read-only.
        /// </summary>
        public static DocumentContent Empty
        {
            get => new DocumentContent();
        }
    }
}
