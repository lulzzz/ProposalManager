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
    public partial class Workflow : BaseArtifact<Workflow>
    {
        /// <summary>
        /// Construcor
        /// </summary>
        public Workflow() : base()
        {
            ContentType = ContentType.Workflow;
            Metadata = new WorkflowMetadata();
            Content = new WorkflowContent();
        }

        /// <summary>
        /// Content type of the opportunity
        /// </summary>
        [JsonProperty("contentType")]
        public new ContentType ContentType { get; private set; }

        /// <summary>
        /// Metadata of the opportunity
        /// </summary>
        [JsonProperty("metadata")]
        public new WorkflowMetadata Metadata { get; set; }

        /// <summary>
        /// Content of the opportunity
        /// </summary>
        [JsonProperty("content")]
        public new WorkflowContent Content { get; set; }

        /// <summary>
        /// Represents the empty opportunity. This field is read-only.
        /// </summary>
        public static Workflow Empty =
            new Workflow();
    }

    public class WorkflowContent
    {
        // Workflows
        // state, etc.
    }

    public class WorkflowMetadata
    {
        // metadata fields
    }
}
