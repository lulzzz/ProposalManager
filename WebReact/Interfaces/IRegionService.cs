// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using ApplicationCore;
using System.Collections.Generic;
using System.Threading.Tasks;
using WebReact.Models;
using WebReact.ViewModels;

namespace WebReact.Interfaces
{
    public interface IRegionService
    {
        Task<StatusCodes> CreateItemAsync(RegionModel modelObject, string requestId = "");

        Task<StatusCodes> UpdateItemAsync(RegionModel modelObject, string requestId = "");

        Task<StatusCodes> DeleteItemAsync(string id, string requestId = "");

        Task<IList<RegionModel>> GetAllAsync(string requestId = "");
    }
}