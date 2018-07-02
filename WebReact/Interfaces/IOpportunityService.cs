// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ApplicationCore;
using ApplicationCore.Artifacts;
using ApplicationCore.Entities;
using Newtonsoft.Json.Linq;
using WebReact.ViewModels;

namespace WebReact.Interfaces
{
    public interface IOpportunityService
    {
        Task<StatusCodes> CreateItemAsync(OpportunityViewModel opportunityViewModel, string requestId = "");
        Task<StatusCodes> DeleteItemAsync(string id, string requestId = "");
        Task<OpportunityViewModel> GetItemByIdAsync(string id, string requestId = "");
        Task<OpportunityViewModel> GetItemByNameAsync(string name, bool isCheckName, string requestId = "");
        Task<OpportunityIndexViewModel> GetAllAsync(int pageIndex, int itemsPage, string requestId = "");
        Task<StatusCodes> UpdateItemAsync(OpportunityViewModel opportunityViewModel, string requestId = "");
        Task<StatusCodes> AddSectionsAsync(string name, IList<DocumentSection> docSections, string requestId = "");
    }
}