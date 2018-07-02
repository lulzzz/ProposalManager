// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ApplicationCore.Artifacts;
using ApplicationCore.Interfaces;
using ApplicationCore.Entities;
using ApplicationCore.Helpers;
using ApplicationCore.Entities.GraphServices;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using Infrastructure.Services;
using ApplicationCore.Helpers.Exceptions;
using System.Linq;

namespace ApplicationCore.Services
{
    public class NotificationRepository : BaseRepository<Notification>, INotificationRepository
    {
        private readonly GraphSharePointAppService _graphSharePointAppService;
        private readonly GraphUserAppService _graphUserAppService;
        private readonly IUserContext _userContext;

        public NotificationRepository(
            ILogger<NotificationRepository> logger, 
            IOptions<AppOptions> appOptions,
            GraphSharePointAppService graphSharePointAppService,
            GraphUserAppService graphUserAppService,
            IUserContext userContext) : base(logger, appOptions)
        {
            Guard.Against.Null(graphSharePointAppService, nameof(graphSharePointAppService));
            Guard.Against.Null(graphUserAppService, nameof(graphUserAppService));
            Guard.Against.Null(userContext, nameof(userContext));
            _graphSharePointAppService = graphSharePointAppService;
            _graphUserAppService = graphUserAppService;
            _userContext = userContext;
        }

        public async Task<StatusCodes> CreateItemAsync(Notification notification, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - NotificationRepository_CreateItemAsync called.");
            
            try
            {
                Guard.Against.Null(notification, "CreateItemAsync_notification null", requestId);
                Guard.Against.NullOrEmpty(notification.Title, "CreateItemAsync_title null-empty", requestId);
                Guard.Against.Null(notification.Fields, "CreateItemAsync_fields null-empty", requestId);
                Guard.Against.NullOrEmpty(notification.Fields.SentTo, "CreateItemAsync_SentTo null-empty", requestId);
                Guard.Against.NullOrEmpty(notification.Fields.Message, "CreateItemAsync_Message null-empty", requestId);
                Guard.Against.NullOrEmpty(notification.Fields.SentFrom, "CreateItemAsync_SentFrom null-empty", requestId);

                // Check result belongs to caller
                var currentUser = (_userContext.User.Claims).ToList().Find(x => x.Type == "preferred_username")?.Value;
                Guard.Against.NullOrEmpty(currentUser, "NotificationRepository_CreateItemAsync CurrentUser null-empty", requestId);

                if (notification.Fields.SentFrom != currentUser)
                {
                    _logger.LogError($"RequestId: {requestId} - NotificationRepository_CreateItemAsync sentFrom: {notification.Fields.SentFrom} current user: {currentUser} AccessDeniedException");
                    throw new AccessDeniedException($"RequestId: {requestId} - NotificationRepository_CreateItemAsync sentFrom: {notification.Fields.SentFrom} current user: {currentUser} AccessDeniedException");
                }

                // Ensure id is blank since it will be set by SharePoint
                notification.Id = String.Empty;

                _logger.LogInformation($"RequestId: {requestId} - NotificationRepository_CreateItemAsync creating SharePoint List for opportunity.");

                // Create Json object for SharePoint update list item
                dynamic notificationFieldsJson = new JObject();
                notificationFieldsJson.Title = notification.Title;
                notificationFieldsJson.SentTo = notification.Fields.SentTo;
                notificationFieldsJson.SentFrom = notification.Fields.SentFrom;
                notificationFieldsJson.SentDate = DateTimeOffset.Now;
                notificationFieldsJson.IsRead = notification.Fields.IsRead;
                notificationFieldsJson.Message = notification.Fields.Message;

                dynamic notificationJson = new JObject();
                notificationJson.fields = notificationFieldsJson;

                var notificationSiteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.NotificationsListId
                };

                var result = await _graphSharePointAppService.CreateListItemAsync(notificationSiteList, notificationJson.ToString(), requestId);

                _logger.LogInformation($"RequestId: {requestId} - NotificationRepository_CreateItemAsync finished creating SharePoint List for notification, sending email now.");

                try
                {
                    var sendEmailResponse = await SendEmailAsync(notification, requestId);
                }
                catch (Exception ex)
                {
                    _logger.LogError($"RequestId: {requestId} - NotificationRepository_CreateItemAsync_SendEmail Service Exception: {ex}");
                }

                return StatusCodes.Status201Created;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - CreateItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - CreateItemAsync Service Exception: {ex}");
            }
        }

        public async Task<Notification> GetItemByIdAsync(string id, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - GetItemByIdAsync called.");

            try
            {
                Guard.Against.NullOrEmpty(id, "GetItemByIdAsync_id null", requestId);

                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.NotificationsListId
                };

                var json = await _graphSharePointAppService.GetListItemByIdAsync(siteList, id, "all", requestId);
                Guard.Against.Null(json, "GetItemByIdAsync_GetListItemByIdAsync json null", requestId);

                var notification = await MapToEntityAsync(json, requestId);

                // Check result belongs to caller
                var currentUser = (_userContext.User.Claims).ToList().Find(x => x.Type == "preferred_username")?.Value;
                Guard.Against.NullOrEmpty(currentUser, "NotificationRepository_GetItemByIdAsync CurrentUser null-empty", requestId);

                if (notification.Fields.SentTo != currentUser)
                {
                    _logger.LogError($"RequestId: {requestId} - NotificationRepository_GetItemByIdAsync id: {id} current user: {currentUser} AccessDeniedException");
                    throw new AccessDeniedException($"RequestId: {requestId} - NotificationRepository_GetItemByIdAsync id: {id} current user: {currentUser} AccessDeniedException");
                }

                return notification;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - GetItemByIdAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - GetItemByIdAsync Service Exception: {ex}");
            }
        }

        public async Task<IList<Notification>> GetAllAsync(string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - GetAllAsync called.");

            try
            {
                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.NotificationsListId
                };

                var currentUser = (_userContext.User.Claims).ToList().Find(x => x.Type == "preferred_username")?.Value;
                Guard.Against.NullOrEmpty(currentUser, "NotificationRepository_GetAllAsync CurrentUser null-empty", requestId);

                var options = new List<QueryParam>();
                options.Add(new QueryParam("filter", $"startswith(fields/SentTo,'{currentUser}')"));

                var json = await _graphSharePointAppService.GetListItemsAsync(siteList, options, "all", requestId);
                JArray jsonArray = JArray.Parse(json["value"].ToString());

                var itemsList = new List<Notification>();
                foreach (var item in jsonArray)
                {
                    itemsList.Add(await MapToEntityAsync(item));
                }
				var sorteditemsList = itemsList.OrderByDescending(x => x.Fields.SentDate).ToList();

				return sorteditemsList;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - GetAllAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - GetAllAsync Service Exception: {ex}");
            }
        }

        public async Task<StatusCodes> UpdateItemAsync(Notification notification, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - UpdateItemAsync called.");
            Guard.Against.Null(notification, "UpdateItemAsync_notification null", requestId);
            Guard.Against.NullOrEmpty(notification.Id, "UpdateItemAsync_id null-empty", requestId);

            try
            {
                // Check result belongs to caller
                var currentUser = (_userContext.User.Claims).ToList().Find(x => x.Type == "preferred_username")?.Value;
                Guard.Against.NullOrEmpty(currentUser, "NotificationRepository_UpdateItemAsync CurrentUser null-empty", requestId);

                if (notification.Fields.SentTo != currentUser)
                {
                    _logger.LogError($"RequestId: {requestId} - NotificationRepository_UpdateItemAsync SentTo: {notification.Fields.SentTo} current user: {currentUser} AccessDeniedException");
                    throw new AccessDeniedException($"RequestId: {requestId} - NotificationRepository_UpdateItemAsync SentTo: {notification.Fields.SentTo} current user: {currentUser} AccessDeniedException");
                }

                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.NotificationsListId
                };

                // Create Json object for SharePoint update list item
                dynamic notificationJson = new JObject();
                notificationJson.IsRead = notification.Fields.IsRead;

                var response = await _graphSharePointAppService.UpdateListItemAsync(siteList, notification.Id, notificationJson.ToString(), requestId);

                _logger.LogInformation($"RequestId: {requestId} - UpdateItemAsync finished SharePoint List for notification.");

                return StatusCodes.Status200OK;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - UpdateItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - UpdateItemAsync Service Exception: {ex}");
            }
        }

        public Task<StatusCodes> DeleteItemAsync(string id, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - DeleteItemAsync called.");
            
            Guard.Against.NullOrEmpty(id, "DeleteItemAsync_id null-empty", requestId);

            try
            {
                // Check result belongs to caller
                var currentUser = (_userContext.User.Claims).ToList().Find(x => x.Type == "preferred_username")?.Value;
                Guard.Against.NullOrEmpty(currentUser, "NotificationRepository_UpdateItemAsync CurrentUser null-empty", requestId);

                //if (notification.Fields.SentTo != currentUser)
                //{
                //    _logger.LogError($"RequestId: {requestId} - NotificationRepository_UpdateItemAsync SentTo: {notification.Fields.SentTo} current user: {currentUser} AccessDeniedException");
                //    throw new AccessDeniedException($"RequestId: {requestId} - NotificationRepository_UpdateItemAsync SentTo: {notification.Fields.SentTo} current user: {currentUser} AccessDeniedException");
                //}

                //var response = await _graphSharePointAppService.

                return Task.FromResult(StatusCodes.Status204NoContent);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - DeleteItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - DeleteItemAsync Service Exception: {ex}");
            }
        }

        // Private methods
        private Task<Notification> MapToEntityAsync(JToken json, string requestId = "")
        {
            try
            {
                var notification = Notification.Empty;
                notification.Id = json["fields"]["id"].ToString() ?? String.Empty;
                notification.Title = json["fields"]["Title"].ToString() ?? String.Empty;
                if (json["fields"]["IsRead"] != null) notification.Fields.IsRead = (bool)json["fields"]["IsRead"];
                notification.Fields.Message = json["fields"]["Message"].ToString() ?? String.Empty;
                if (json["fields"]["SentDate"] != null) notification.Fields.SentDate = DateTimeOffset.Parse(json["fields"]["SentDate"].ToString());
                notification.Fields.SentTo = json["fields"]["SentTo"].ToString() ?? String.Empty;
                notification.Fields.SentFrom = json["fields"]["SentFrom"].ToString() ?? String.Empty;

                return Task.FromResult<Notification>(notification);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - MapToEntity Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - MapToEntity Service Exception: {ex}");
            }

            
        }

        private async Task<JObject> SendEmailAsync(Notification notification, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - NotificationRepository_SendEmail start");

            try
            {
                // Create Json object for email message
                dynamic bodyJson = new JObject();
                bodyJson.contentType = "Text";
                bodyJson.content = notification.Fields.Message;

                dynamic recipientsJson = new JObject();
                recipientsJson.address = notification.Fields.SentTo;

                dynamic toRecipientsJson = new JObject();
                toRecipientsJson.emailAddress = recipientsJson;

                var array = new JArray();
                array.Add(toRecipientsJson);

                dynamic messageJson = new JObject();
                messageJson.subject = notification.Title;
                messageJson.body = bodyJson;
                messageJson.toRecipients = array;

                dynamic requestJson = new JObject();
                requestJson.message = messageJson;


                var fromEmail = notification.Fields.SentFrom;
                if (!String.IsNullOrEmpty(_appOptions.ServiceEmail))
                {
                    fromEmail = _appOptions.ServiceEmail;
                }

                var response = await _graphUserAppService.SendEmail(fromEmail, requestJson.ToString(), requestId);

                _logger.LogInformation($"RequestId: {requestId} - NotificationRepository_SendEmail sent completed");
                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - NotificationRepository_SendEmail Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - NotificationRepository_SendEmail Service Exception: {ex}");
            }
        }
    }
}
