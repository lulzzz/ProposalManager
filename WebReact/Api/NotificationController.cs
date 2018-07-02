// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ApplicationCore;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using WebReact.Interfaces;
using ApplicationCore.Helpers;
using ApplicationCore.Artifacts;
using Newtonsoft.Json.Linq;
using WebReact.Models;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Authorization;

namespace WebReact.Api
{
    public class NotificationController : BaseApiController<NotificationController>
    {
        private readonly INotificationService _notificationService;

        public NotificationController(
            ILogger<NotificationController> logger, 
            IOptions<AppOptions> appOptions, 
            INotificationService notificationService) : base(logger, appOptions)
        {
            Guard.Against.Null(notificationService, nameof(notificationService));
            _notificationService = notificationService;
        }

        // POST: /Notification
        [Authorize]
        [HttpPost]
        public async Task<IActionResult> Create([FromBody] JObject notificationJson)
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"RequestID:{requestId} - Create called.");

            try
            {
                if (notificationJson == null)
                {
                    _logger.LogError($"RequestID:{requestId} - Create error: notification null");
                    var errorResponse = JsonErrorResponse.BadRequest($"Create error: notification null", requestId);

                    return BadRequest(errorResponse);
                }

                var notification = JsonConvert.DeserializeObject<NotificationModel>(notificationJson.ToString(), new JsonSerializerSettings
                {
                    MissingMemberHandling = MissingMemberHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore
                });

                if (String.IsNullOrEmpty(notification.Title) || String.IsNullOrEmpty(notification.Message) || String.IsNullOrEmpty(notification.SentFrom) || String.IsNullOrEmpty(notification.SentTo))
                {
                    _logger.LogError($"RequestID:{requestId} - Create error: invalid parameters");
                    var errorResponse = JsonErrorResponse.BadRequest($"Create error: invalid parameters", requestId);

                    return BadRequest(errorResponse);
                }

                var resultCode = await _notificationService.CreateItemAsync(notification, requestId);

                if (resultCode != ApplicationCore.StatusCodes.Status201Created)
                {
                    _logger.LogError($"RequestID:{requestId} - Create error: {resultCode.Name}");
                    var errorResponse = JsonErrorResponse.BadRequest($"Create error: notification {resultCode.Name}", requestId);

                    return BadRequest(errorResponse);
                }

                var location = "/Notification/Create/new"; // TODO: Get the id from the results but need to wire from factory to here

                return Created(location, $"RequestId: {requestId} - Notification created.");
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestID:{requestId} Create error: {ex.Message}");
                var errorResponse = JsonErrorResponse.BadRequest($"Create error: {ex} ", requestId);

                return BadRequest(errorResponse);
            }
        }

        // PUT: /Notification?id={id}
        [Authorize]
        [HttpPatch]
        public async Task<IActionResult> Update([FromBody] JObject notificationJson)
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"RequestID:{requestId} - Update called.");

            try
            {
                if (notificationJson == null)
                {
                    _logger.LogError($"RequestID:{requestId} - Update error: notification null");
                    var errorResponse = JsonErrorResponse.BadRequest($"Update error: notification null", requestId);

                    return BadRequest(errorResponse);
                }

                var notification = JsonConvert.DeserializeObject<NotificationModel>(notificationJson.ToString(), new JsonSerializerSettings
                {
                    MissingMemberHandling = MissingMemberHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore
                });

                if (String.IsNullOrEmpty(notification.Id) || String.IsNullOrEmpty(notification.SentFrom))
                {
                    _logger.LogError($"RequestID:{requestId} - Update error: invalid notification id");
                    var errorResponse = JsonErrorResponse.BadRequest($"Update error: invalid notification id", requestId);

                    return BadRequest(errorResponse);
                }

                var resultCode = await _notificationService.UpdateItemAsync(notification, requestId);

                if (resultCode != ApplicationCore.StatusCodes.Status200OK)
                {
                    _logger.LogError($"RequestID:{requestId} - Update error: {resultCode.Name}");
                    var errorResponse = JsonErrorResponse.BadRequest($"Update error: {resultCode.Name} ", requestId);

                    return BadRequest(errorResponse);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestID:{requestId} - Update error: {ex.Message}");
                var errorResponse = JsonErrorResponse.BadRequest($"Update error: {ex.Message} ", requestId);

                return BadRequest(errorResponse);
            }
        }

        // DELETE: /Notification?id={id}
        [Authorize]
        [HttpDelete("{id}")]
        public async Task<IActionResult> Delete(string id)
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"RequestID:{requestId} - Delete called.");

            try
            {
                if (String.IsNullOrEmpty(id))
                {
                    _logger.LogError($"RequestID:{requestId} - Delete error: id null");
                    var errorResponse = JsonErrorResponse.BadRequest($"Delete error: Delete id null", requestId);

                    return BadRequest(errorResponse);
                }

                var resultCode = await _notificationService.DeleteItemAsync(id, requestId);

                if (resultCode != ApplicationCore.StatusCodes.Status200OK)
                {
                    _logger.LogError($"RequestID:{requestId} - Delete error: {resultCode.Name}");
                    var errorResponse = JsonErrorResponse.BadRequest($"Delete error: {resultCode.Name} ", requestId);

                    return BadRequest(errorResponse);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestID:{requestId} - Delete error: {ex.Message}");
                var errorResponse = JsonErrorResponse.BadRequest($"Delete error: {ex.Message} ", requestId);

                return BadRequest(errorResponse);
            }
        }

        // GET: /Notification
        [Authorize]
        [HttpGet]
        public async Task<IActionResult> GetAll(int? page, [FromQuery] string id)
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"RequestID:{requestId} - GetAll called.");

            try
            {
                if (!String.IsNullOrEmpty(id))
                {
                    return await GetById(id);
                }

                var modelList = (await _notificationService.GetAllAsync(1, 10, requestId)).ToList();
                Guard.Against.Null(modelList, nameof(modelList), requestId);
                if (modelList.Count == 0)
                {
                    _logger.LogError($"RequestID:{requestId} - GetAll no items found.");
                    return NotFound($"RequestID:{requestId} - GetAll no items found");
                }

                return Ok(modelList);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestID:{requestId} - GetAll error: {ex.Message}");
                var errorResponse = JsonErrorResponse.BadRequest($"GetAll error: {ex.Message} ", requestId);

                return BadRequest(errorResponse);
            }
        }

        // GET: /Notification?id={id}
        public async Task<IActionResult> GetById(string id)
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"RequestID:{requestId} - GetById called.");

            try
            {
                if (String.IsNullOrEmpty(id))
                {
                    _logger.LogError($"RequestID:{requestId} - GetById notification id == null.");
                    return NotFound($"RequestID:{requestId} - GetById notification Invalid parameter. id = null ");
                }
                var response = await _notificationService.GetItemByIdAsync(id, requestId);
                if (response == null)
                {
                    _logger.LogError($"RequestID:{requestId} - GetById no notifications found.");
                    return NotFound($"RequestID:{requestId} - GetById no notifications found");
                }

                var responseJObject = JObject.FromObject(response);

                return Ok(responseJObject);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestID:{requestId} - GetById error: {ex.Message}");
                var errorResponse = JsonErrorResponse.BadRequest($"GetById error: {ex.Message} ", requestId);

                return BadRequest(errorResponse);
            }
        }
    }
}
