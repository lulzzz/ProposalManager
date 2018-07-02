// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using ApplicationCore.Entities;
using System.Collections.Generic;
using System.Threading.Tasks;
using WebReact.Models;
using WebReact.ViewModels;

namespace WebReact.Interfaces
{
    public interface IUserProfileService
    {
        Task<UserProfileViewModel> GetItemByIdAsync(string id, string requestId = "");
        Task<UserProfileListViewModel> GetAllAsync(int pageIndex, int itemsPage, string requestId = "");
		//Task<UserProfileViewModel> GetByUpn(string upn);
		Task<UserProfileViewModel> GetItemByUpnAsync(string upn, string requestId = "");
    }
}