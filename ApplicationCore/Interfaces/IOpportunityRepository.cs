// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using ApplicationCore.Artifacts;
using Newtonsoft.Json.Linq;

namespace ApplicationCore.Interfaces
{
    public interface IOpportunityRepository : IArtifactFactory<Opportunity>
    {
        Task<StatusCodes> CreateItemAsync(Opportunity opportunity, string requestId = "");
        Task<StatusCodes> DeleteItemAsync(string id, string requestId = "");
        Task<Opportunity> GetItemByIdAsync(string id, string requestId = "");
        Task<Opportunity> GetItemByNameAsync(string name, bool isCheckName, string requestId = "");
        Task<IList<Opportunity>> GetAllAsync(string requestId = "");
        Task<StatusCodes> UpdateItemAsync(Opportunity opportunity, string requestId = "");
    }
}
