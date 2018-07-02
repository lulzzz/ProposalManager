// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebReact.Models
{
    public class ProposalDocumentModel : DocumentModel
    {
        [JsonProperty("content")]
        public new ProposalDocumentContentModel Content { get; set; }
    }

    public class ProposalDocumentContentModel
    {
        /// <summary>
        /// Proposal document sections list
        /// </summary>
        [JsonProperty("proposalSectionList")]
        public IList<DocumentSectionModel> ProposalSectionList { get; set; }
    }
}
