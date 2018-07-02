// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System.Collections.Generic;
using System.Threading.Tasks;
using ApplicationCore.Artifacts;
using ApplicationCore.Entities;

namespace ApplicationCore.Interfaces
{
    public interface INotificationRepository
    {
        Task<StatusCodes> CreateItemAsync(Notification notification, string requestId = "");
        Task<StatusCodes> DeleteItemAsync(string id, string requestId = "");
        Task<Notification> GetItemByIdAsync(string id, string requestId = "");
        Task<IList<Notification>> GetAllAsync(string requestId = "");
        Task<StatusCodes> UpdateItemAsync(Notification notification, string requestId = "");
    }
}