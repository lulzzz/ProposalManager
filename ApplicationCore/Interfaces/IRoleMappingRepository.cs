// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System.Collections.Generic;
using System.Threading.Tasks;
using ApplicationCore.Entities;

namespace ApplicationCore.Interfaces
{
    public interface IRoleMappingRepository
    {
        Task<StatusCodes> CreateItemAsync(RoleMapping entity, string requestId = "");

        Task<StatusCodes> UpdateItemAsync(RoleMapping entity, string requestId = "");

        Task<StatusCodes> DeleteItemAsync(string id, string requestId = "");

        Task<IList<RoleMapping>> GetAllAsync(string requestId = "");
    }
}
