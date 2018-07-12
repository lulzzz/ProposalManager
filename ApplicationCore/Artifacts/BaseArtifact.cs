// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace ApplicationCore.Artifacts
{
    /// <summary>
    /// Base artifact for data and object composition
    /// </summary>
    public abstract class BaseArtifact
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public BaseArtifact()
        {
            ContentType = ContentType.NoneEmpty;
        }

        /// <summary>
        /// Unique identifier of the artifact
        /// </summary>
        [JsonProperty("id", Order = 1)]
        public string Id { get; set; }

        /// <summary>
        /// Display name of the artifact
        /// </summary>
        [JsonProperty("displayName", Order = 2)]
        public string DisplayName { get; set; }

        /// <summary>
        /// Internal reference id for the artifact used to link the artifact in external systems
        /// </summary>
        [JsonProperty("reference", Order = 3)]
        public string Reference { get; set; }

        /// <summary>
        /// Content type of the artifact
        /// </summary>
        [JsonProperty("contentType", Order = 4)]
        public ContentType ContentType { get; set; }

        /// <summary>
        /// Version of the artifact
        /// </summary>
        [JsonProperty("version", Order = 5)]
        public string Version { get; set; }

        /// <summary>
        /// Artifact Uri composed as: {ContentType}://{Id}/{Name}/{Version}
        /// </summary>
        [JsonProperty("uri", Order = 6)]
        public string Uri { get { return ContentType.Name + "://" + Id + "/" + Version ?? "1.0"; } }
    }

    public partial class BaseArtifact<T> : BaseArtifact
    {
        /// <summary>
        /// Artifact type name
        /// </summary>
        [JsonProperty("typeName", Order = 7)]
        public string TypeName { get { return typeof(T).Name; } }

        /// <summary>
        /// Metadata of the artifact
        /// </summary>
        [JsonProperty("metadata", Order = 8)]
        public T Metadata { get; set; }

        /// <summary>
        /// Content of the artifact
        /// </summary>
        [JsonProperty("content", Order = 9)]
        public T Content { get; set; }
    }
}
