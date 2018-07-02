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
    public interface INotificationService
    {
        Task<StatusCodes> CreateItemAsync(NotificationModel notificationModel, string requestId = "");
        Task<StatusCodes> DeleteItemAsync(string id, string requestId = "");
        Task<IList<NotificationModel>> GetAllAsync(int pageIndex, int itemsPage, string requestId = "");
        Task<NotificationModel> GetItemByIdAsync(string id, string requestId = "");
        Task<StatusCodes> UpdateItemAsync(NotificationModel notificationModel, string requestId = "");
    }
}