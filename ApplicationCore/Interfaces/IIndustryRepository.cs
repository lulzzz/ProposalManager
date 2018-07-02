// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System.Collections.Generic;
using System.Threading.Tasks;
using ApplicationCore.Entities;

namespace ApplicationCore.Interfaces
{
    public interface IIndustryRepository
    {
        Task<StatusCodes> CreateItemAsync(Industry entity, string requestId = "");

        Task<StatusCodes> UpdateItemAsync(Industry entity, string requestId = "");

        Task<StatusCodes> DeleteItemAsync(string id, string requestId = "");

        Task<IList<Industry>> GetAllAsync(string requestId = "");
    }
}