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
using Microsoft.AspNetCore.Authorization;
using System.IO;
using ApplicationCore.Interfaces;

namespace WebReact.Api
{
    public class ContextController : BaseApiController<ContextController>
    {
        private readonly IContextService _contextService;
        private readonly IOpportunityService _opportunityService;


        public ContextController(
            ILogger<ContextController> logger, 
            IOptions<AppOptions> appOptions,
            IContextService contextService,
            IOpportunityService opportunityService) : base(logger, appOptions)
        {
            Guard.Against.Null(contextService, nameof(contextService));
            Guard.Against.Null(opportunityService, nameof(opportunityService));
            _contextService = contextService;
            _opportunityService = opportunityService;
        }

        // Get: /Context/GetSiteDrive
        [HttpGet("GetSiteDrive/{siteName}", Name = "GetSiteDrive")]
        public async Task<IActionResult> GetSiteDrive(string siteName)
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"RequestID:{requestId} GetSiteDrive called.");

            try
            {
                if (siteName == null)
                {
                    _logger.LogError($"RequestID:{requestId} GetSiteDrive error: siteName null");
                    var errorResponse = JsonErrorResponse.BadRequest("GetSiteDrive error: siteName null", requestId);

                    return BadRequest(errorResponse);
                }

                var response = await _contextService.GetSiteDriveAsync(siteName);

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestID:{requestId} GetSiteDrive error: {ex.Message}");
                var errorResponse = JsonErrorResponse.BadRequest($"RequestID:{requestId} GetSiteDrive error: {ex.Message}", requestId);

                return BadRequest(errorResponse);
            }
        }

        // Get: /Context/GetSiteDrive
        [HttpGet("GetSiteId/{siteName}", Name = "GetSiteId")]
        public async Task<IActionResult> GetSiteId(string siteName)
        {
            var requestId = Guid.NewGuid().ToString();
            _logger.LogInformation($"RequestID:{requestId} GetSiteId called.");

            try
            {
                if (siteName == null)
                {
                    _logger.LogError($"RequestID:{requestId} GetSiteId error: siteName null");
                    var errorResponse = JsonErrorResponse.BadRequest("GetSiteId error: siteName null", requestId);

                    return BadRequest(errorResponse);
                }

                var response = await _contextService.GetSiteIdAsync(siteName);

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestID:{requestId} GetSiteId error: {ex.Message}");
                var errorResponse = JsonErrorResponse.BadRequest($"GetSiteId error: {ex.Message}", requestId);

                return BadRequest(errorResponse);
            }
        }
    }
}
