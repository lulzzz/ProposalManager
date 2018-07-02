// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using WebReact.ViewModels;
using WebReact.Interfaces;
using ApplicationCore;
using ApplicationCore.Artifacts;
using Infrastructure.Services;
using ApplicationCore.Interfaces;
using ApplicationCore.Helpers;
using WebReact.Models;
using ApplicationCore.Entities;
using System.ComponentModel;
using System.Globalization;
using ApplicationCore.Helpers.Exceptions;

namespace WebReact.Services
{
    public class NotificationService : BaseService<NotificationService>, INotificationService
    {
        private readonly INotificationRepository _notificationRepository;

        public NotificationService(
            ILogger<NotificationService> logger,
            IOptions<AppOptions> appOptions,
            INotificationRepository notificationRepository) : base(logger, appOptions)
        {
            Guard.Against.Null(notificationRepository, nameof(notificationRepository));
            _notificationRepository = notificationRepository;
        }

        public async Task<IList<NotificationModel>> GetAllAsync(int pageIndex, int itemsPage, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - GetAllAsync called.");

            try
            {
                var listItems = (await _notificationRepository.GetAllAsync(requestId)).ToList();
                Guard.Against.Null(listItems, "GetAllAsync_listItems null", requestId);

                var modelListItems = new List<NotificationModel>();
                foreach (var item in listItems)
                {
                    modelListItems.Add(await MapToModelAsync(item, requestId));
                }

                if (modelListItems.Count == 0)
                {
                    _logger.LogWarning($"RequestId: {requestId} - GetAllAsync no items found");
                    throw new NoItemsFound($"RequestId: {requestId} - Method name: GetAllAsync - No Items Found");
                }

                return modelListItems;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - GetAllAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - GetAllAsync Service Exception: {ex}");
            }
        }

        public async Task<NotificationModel> GetItemByIdAsync(string id, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - GetItemByIdAsync called.");

            try
            {
                Guard.Against.NullOrEmpty(id, "GetItemByIdAsync_id null", requestId);

                var notification = await _notificationRepository.GetItemByIdAsync(id, requestId);
                Guard.Against.Null(notification, "GetItemByIdAsync_notification null", requestId);

                var notificationModel = await MapToModelAsync(notification, requestId);

                return notificationModel;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - GetItemByIdAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - GetItemByIdAsync Service Exception: {ex}");
            }
        }

        public async Task<StatusCodes> DeleteItemAsync(string id, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - DeleteItemAsync called.");

            try
            {
                Guard.Against.NullOrEmpty(id, "DeleteItemAsync_id null", requestId);

                return await _notificationRepository.DeleteItemAsync(id, requestId);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - DeleteItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - DeleteItemAsync Service Exception: {ex}");
            }
        }

        public async Task<StatusCodes> UpdateItemAsync(NotificationModel notificationModel, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - UpdateItemAsync called.");

            Guard.Against.Null(notificationModel, nameof(notificationModel), requestId);
            Guard.Against.NullOrEmpty(notificationModel.Id, nameof(notificationModel.Id), requestId);

            try
            {
                Guard.Against.Null(notificationModel, "UpdateItemAsync_notificationModel null", requestId);

                var entity = await MapToEntityAsync(notificationModel, requestId);

                var result = await _notificationRepository.UpdateItemAsync(entity, requestId);

                Guard.Against.NotStatus200OK(result, "UpdateItemAsync", requestId);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - UpdateItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - UpdateItemAsync Service Exception: {ex}");
            }
        }

        public async Task<StatusCodes> CreateItemAsync(NotificationModel notificationModel, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - UpdateItemAsync called.");

            try
            {
                Guard.Against.Null(notificationModel, "CreateItemAsync_notificationModel null", requestId);

                var entity = await MapToEntityAsync(notificationModel, requestId);

                var result = await _notificationRepository.CreateItemAsync(entity, requestId);

                Guard.Against.NotStatus201Created(result, "CreateItemAsync", requestId);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - CreateItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - CreateItemAsync Service Exception: {ex}");
            }
        }


        // Private methods
        private Task<NotificationModel> MapToModelAsync(Notification entity, string requestId = "")
        {
            try
            {
                var dto = new NotificationModel
                {
                    Id = entity.Id,
                    Title = entity.Title,
                    IsRead = entity.Fields.IsRead,
                    Message = entity.Fields.Message,
                    SentDate = entity.Fields.SentDate,
                    SentFrom = entity.Fields.SentFrom,
                    SentTo = entity.Fields.SentTo
                };

                return Task.FromResult(dto);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - MapToModel Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - MapToModel Service Exception: {ex}");
            }
        }

        private Task<Notification> MapToEntityAsync(NotificationModel model, string requestId = "")
        {
            try
            {
                var entity = new Notification
                {
                    Id = model.Id,
                    Fields = new NotificationFields
                    {
                        IsRead = model.IsRead,
                        Message = model.Message,
                        SentDate = model.SentDate,
                        SentFrom = model.SentFrom,
                        SentTo = model.SentTo
                    },
                    Title = model.Title
                };

                return Task.FromResult(entity);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - MapToEntity Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - MapToEntity Service Exception: {ex}");
            }
        }
    }
}
