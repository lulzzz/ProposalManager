// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System.Collections.Generic;
using System.Threading.Tasks;
using ApplicationCore.Entities;

namespace ApplicationCore.Interfaces
{
    public interface IUserProfileRepository
    {
        Task<UserProfile> GetItemByIdAsync(string id, string requestId = "");
        Task<IList<UserProfile>> GetAllAsync(string requestId = "");
		Task<UserProfile> GetItemByUpnAsync(string upn, string requestId = "");
	}
}