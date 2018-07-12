// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace ApplicationCore.Artifacts.Workflows
{
    public class CreateOpportunityWorkflow : BaseArtifact<CreateOpportunityWorkflow>
    {
        /// <summary>
        /// Construcor
        /// </summary>
        public CreateOpportunityWorkflow() : base()
        {
            Id = Guid.NewGuid().ToString();
            ContentType = ContentType.Workflow;
            Metadata = new CreateOpportunityWorkflowMetadata();
            Content = new CreateOpportunityWorkflowContent();
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
        public new CreateOpportunityWorkflowMetadata Metadata { get; set; }

        /// <summary>
        /// Content of the opportunity
        /// </summary>
        [JsonProperty("content")]
        public new CreateOpportunityWorkflowContent Content { get; set; }

        /// <summary>
        /// Represents the empty opportunity. This field is read-only.
        /// </summary>
        public static CreateOpportunityWorkflow Empty()
        {
            return new CreateOpportunityWorkflow();
        }
    }

    public class CreateOpportunityWorkflowContent
    {
        // Workflows
        // state, etc.
    }

    public class CreateOpportunityWorkflowMetadata
    {
        public CreateOpportunityWorkflowMetadata()
        {
            TriggerUri = @"https://prod-23.westus.logic.azure.com:443/workflows/2132c8e4a349414680a074429f003f30/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-TnYFpFiRSzX87ZEQ4PJbaPNvxNvUeFmKSMFZ58JJQs";
        }

        public string TriggerUri { get; }
    }
}
