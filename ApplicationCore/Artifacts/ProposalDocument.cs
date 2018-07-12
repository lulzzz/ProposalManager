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
    public class ProposalDocument : Document
    {
        public ProposalDocument() : base()
        {
            ContentType = ContentType.ProposalDocument;
        }

        /// <summary>
        /// Content of the document
        /// </summary>
        [JsonProperty("content")]
        public new ProposalDocumentContent Content { get; set; }

        /// <summary>
        /// Represents the empty proposal document. This field is read-only.
        /// </summary>
        public static new ProposalDocument Empty
        {
            get => new ProposalDocument
            {
                Id = String.Empty,
                DisplayName = String.Empty,
                Reference = String.Empty,
                ContentType = ContentType.Opportunity,
                Version = "1.0",
                Metadata = DocumentMetadata.Empty,
                Content = ProposalDocumentContent.Empty
            };
        } 
    }

    public class ProposalDocumentContent
    {
        /// <summary>
        /// Proposal document sections list
        /// </summary>
        [JsonProperty("proposalSectionList")]
        public IList<DocumentSection> ProposalSectionList { get; set; }

        /// <summary>
        /// Represents the empty proposal document. This field is read-only.
        /// </summary>
        public static ProposalDocumentContent Empty
        {
            get => new ProposalDocumentContent
            {
                ProposalSectionList = new List<DocumentSection>()
            };
        }
    }
}
