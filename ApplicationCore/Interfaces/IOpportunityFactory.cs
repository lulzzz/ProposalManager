// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using ApplicationCore.Artifacts;
using ApplicationCore.Entities;
using Newtonsoft.Json.Linq;

namespace ApplicationCore.Interfaces
{
    public interface IOpportunityFactory : IArtifactFactory<Opportunity>
    {
        Task<bool> CheckAccessAsync(Opportunity oppArtifact, List<Role> roles, string requestId = "");

        Task<bool> CheckAccessAnyAsync(Opportunity oppArtifact, string requestId = "");

        Task<Opportunity> CreateWorkflowAsync(Opportunity opportunity, string requestId = "");

        Task<Opportunity> UpdateWorkflowAsync(Opportunity opportunity, string requestId = "");

        Task<Opportunity> MoveTempFileToTeamAsync(Opportunity opportunity, string requestId = "");

        Task<IList<Checklist>> RemoveEmptyFromChecklistAsync(IList<Checklist> checklists, string requestId = "");
    }
}
