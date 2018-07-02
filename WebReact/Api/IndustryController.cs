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
using WebReact.ViewModels;
using WebReact.Models;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Authorization;

namespace WebReact.Api
{
    public class IndustryController : BaseApiController<IndustryController>
    {
        private readonly IIndustryService _industryService;

        public IndustryController(
            ILogger<IndustryController> logger, 
            IOptions<AppOptions> appOptions,
            IIndustryService industryService) : base(logger, appOptions)
        {
            Guard.Against.Null(industryService, nameof(industryService));
            _industryService = industryService;
        }

        [Authorize]
        [HttpPost]
        public async Task<IActionResult> Create([FromBody] JObject jsonObject)
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"RequestID:{requestId} - Industry_Create called.");

            try
            {
                if (jsonObject == null)
                {
                    _logger.LogError($"RequestID:{requestId} - Industry_Create error: null");
                    var errorResponse = JsonErrorResponse.BadRequest($"Industry_Create error: null", requestId);

                    return BadRequest(errorResponse);
                }

                var modelObject = JsonConvert.DeserializeObject<IndustryModel>(jsonObject.ToString(), new JsonSerializerSettings
                {
                    MissingMemberHandling = MissingMemberHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore
                });

                //TODO: P2 Refactor into Guard
                if (String.IsNullOrEmpty(modelObject.Name))
                {
                    _logger.LogError($"RequestID:{requestId} - Industry_Create error: invalid name");
                    var errorResponse = JsonErrorResponse.BadRequest($"Industry_Create error: invalid name", requestId);

                    return BadRequest(errorResponse);
                }

                var resultCode = await _industryService.CreateItemAsync(modelObject, requestId);

                if (resultCode != ApplicationCore.StatusCodes.Status201Created)
                {
                    _logger.LogError($"RequestID:{requestId} - Industry_Create error: {resultCode.Name}");
                    var errorResponse = JsonErrorResponse.BadRequest($"Industry_Create error: {resultCode.Name}", requestId);

                    return BadRequest(errorResponse);
                }

                var location = "/Industry/Create/new"; // TODO: Get the id from the results but need to wire from factory to here

                return Created(location, $"RequestId: {requestId} - Industry created.");
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestID:{requestId} Industry_Create error: {ex.Message}");
                var errorResponse = JsonErrorResponse.BadRequest($"Industry_Create error: {ex} ", requestId);

                return BadRequest(errorResponse);
            }
        }

        [Authorize]
        [HttpPatch]
        public async Task<IActionResult> Update([FromBody] JObject jsonObject)
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"RequestID:{requestId} - Industry_Update called.");

            try
            {
                if (jsonObject == null)
                {
                    _logger.LogError($"RequestID:{requestId} - Industry_Update error: null");
                    var errorResponse = JsonErrorResponse.BadRequest($"Industry_Update error: null", requestId);

                    return BadRequest(errorResponse);
                }

                var modelObject = JsonConvert.DeserializeObject<IndustryModel>(jsonObject.ToString(), new JsonSerializerSettings
                {
                    MissingMemberHandling = MissingMemberHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore
                });

                //TODO: P2 Refactor into Guard
                if (String.IsNullOrEmpty(modelObject.Id))
                {
                    _logger.LogError($"RequestID:{requestId} - Industry_Update error: invalid id");
                    var errorResponse = JsonErrorResponse.BadRequest($"Industry_Update error: invalid id", requestId);

                    return BadRequest(errorResponse);
                }

                var resultCode = await _industryService.UpdateItemAsync(modelObject, requestId);

                if (resultCode != ApplicationCore.StatusCodes.Status200OK)
                {
                    _logger.LogError($"RequestID:{requestId} - Industry_Update error: {resultCode.Name}");
                    var errorResponse = JsonErrorResponse.BadRequest($"Industry_Update error: {resultCode.Name} ", requestId);

                    return BadRequest(errorResponse);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestID:{requestId} - Industry_Update error: {ex.Message}");
                var errorResponse = JsonErrorResponse.BadRequest($"Industry_Update error: {ex.Message} ", requestId);

                return BadRequest(errorResponse);
            }
        }

        [Authorize]
        [HttpDelete("{id}")]
        public async Task<IActionResult> Delete(string id)
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"RequestID:{requestId} - Industry_Delete called.");

            if (String.IsNullOrEmpty(id))
            {
                _logger.LogError($"RequestID:{requestId} - Industry_Delete id == null.");
                return NotFound($"RequestID:{requestId} - Industry_Delete Null ID passed");
            }

            var resultCode = await _industryService.DeleteItemAsync(id, requestId);

            if (resultCode != ApplicationCore.StatusCodes.Status204NoContent)
            {
                _logger.LogError($"RequestID:{requestId} - Industry_Delete error: " + resultCode);
                var errorResponse = JsonErrorResponse.BadRequest($"Industry_Delete error: {resultCode.Name} ", requestId);

                return BadRequest(errorResponse);
            }

            return NoContent();
        }

        [Authorize]
        [HttpGet]
        public async Task<IActionResult> GetAll()
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"RequestID:{requestId} - GetAll called.");

            try
            {
                var modelList = (await _industryService.GetAllAsync(requestId)).ToList();
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
    }
}
