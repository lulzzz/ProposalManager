// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace ApplicationCore.Entities
{
    /// <summary>
    /// Provides a base class for entities, the root object contains the metadata and fields the data
    /// </summary>
    public abstract class BaseEntity
    {
        /// <summary>
        /// Entity identifier
        /// </summary>
        /// <value>Unique ID to identify the model data</value>
        [JsonProperty("id", Order = 1)]
        public string Id { get; set; }
    }

    [JsonObject(MemberSerialization.OptIn)]
    public abstract class BaseEntity<T> : BaseEntity
    {
        /// <summary>
        /// Entity type name
        /// </summary>
        /// <value>Type name for this entity</value>
        [JsonProperty("typeName")]
        public string TypeName { get { return typeof(T).Name; } }

        /// <summary>
        /// The values representing the entity itself
        /// </summary>
        //[JsonProperty("fields")]
        //public T Fields { get; set; }
    }
}
