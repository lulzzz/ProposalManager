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

namespace WebReact.Api
{
    public class UserProfileController : BaseApiController<UserProfileController>
    {
        private readonly IUserProfileService _userProfileService;

        public UserProfileController(
            ILogger<UserProfileController> logger, 
            IOptions<AppOptions> appOptions,
            IUserProfileService userProfileService) : base(logger, appOptions)
        {
            Guard.Against.Null(userProfileService, nameof(userProfileService));
            _userProfileService = userProfileService;
        }

        // GET: /UserProfile/me
        [HttpGet("me", Name = "GetUserProfile")]
        public async Task<IActionResult> GetById()
        {
            _logger.LogInformation("GetById called.");

            // TODO: get id from user context
            var id = "usercontext";

            var thisUserProfile = await _userProfileService.GetItemByIdAsync(id);
            var responseJObject = JObject.FromObject(thisUserProfile);
            return Ok(responseJObject);
        }

        // GET: /UserProfile
        [HttpGet]
        public async Task<IActionResult> GetAll(int? page, [FromQuery] string upn)
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"RequestID:{requestId} - GetAll called.");

            try
            {
                if (!String.IsNullOrEmpty(upn))
                {
					return await GetByUpn(upn);
					
				}

                var itemsPage = 10;
                var modelList = await _userProfileService.GetAllAsync(page ?? 1, itemsPage, requestId);
                Guard.Against.Null(modelList, nameof(modelList), requestId);

                if (modelList.ItemsList.Count == 0)
                {
                    _logger.LogError($"RequestID:{requestId} - GetAll no items found.");
                    return NotFound($"RequestID:{requestId} - GetAll no items found");
                }

                var responseJson = JObject.FromObject(modelList);

                return Ok(responseJson);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestID:{requestId} - GetAll error: {ex.Message}");
                var errorResponse = JsonErrorResponse.BadRequest($"GetAll error: {ex.Message} ", requestId);

                return BadRequest(errorResponse);
            }
        }

        // GET: /UserProfile?upn={name}
        //[Authorize]
        //[HttpGet("name", Name = "GetOpportunityByName")]
        public async Task<IActionResult> GetByUpn(string upn)
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"UPN:{upn} - GetByUPN called.");

            try
            {
				if (String.IsNullOrEmpty(upn))
                {
                    _logger.LogError($"UPN:{requestId} - GetByUPN name == null.");
                    return NotFound($"UPN:{requestId} - GetByUPN Invalid parameter passed");
                }
                var userProfile = await _userProfileService.GetItemByUpnAsync(upn);
                if (userProfile == null)
                {
                    _logger.LogError($"UPN:{requestId} - GetByUPN no user found.");
                    return NotFound($"UPN:{requestId} - GetByUPN no user found");
                }

                var responseJObject = JObject.FromObject(userProfile);

                return Ok(responseJObject);
            }
            catch (Exception ex)
            {
                _logger.LogError($"UPN:{requestId} - GetByUPN error: {ex.Message}");
                var errorResponse = JsonErrorResponse.BadRequest($"GetByUPN error: {ex.Message} ", requestId);

                return BadRequest(errorResponse);
            }
        }
    }
}
