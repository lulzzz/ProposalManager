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
using Newtonsoft.Json.Linq;

namespace WebReact.Services
{
    public class ContextService : BaseService<ContextService>, IContextService
    {
        private readonly GraphSharePointAppService _graphSharePointAppService;

        public ContextService(
            ILogger<ContextService> logger, 
            IOptions<AppOptions> appOptions,
            GraphSharePointAppService graphSharePointAppService) : base(logger, appOptions)
        {
            Guard.Against.Null(graphSharePointAppService, nameof(graphSharePointAppService));
            _graphSharePointAppService = graphSharePointAppService;
        }

        public async Task<JObject> GetTeamGroupDriveAsync(string teamGroupName)
        {
            _logger.LogInformation("GetTeamGroupDriveAsync called.");

            try
            {
                Guard.Against.NullOrEmpty(teamGroupName, nameof(teamGroupName));
                string result = string.Concat(teamGroupName.Where(c => !char.IsWhiteSpace(c)));

                // TODO: Implement,, the below code is part of boilerplate
                var siteIdResponse = await _graphSharePointAppService.GetSiteIdAsync(_appOptions.SharePointHostName, result);
                dynamic responseDyn = siteIdResponse;
                var siteId = responseDyn.id.ToString();

                var driveResponse = await _graphSharePointAppService.GetSiteDriveAsync(siteId);

                return driveResponse;
            }
            catch (Exception ex)
            {
                _logger.LogError("GetTeamGroupDriveAsync error: " + ex);
                throw;
            }
            
        }

        public async Task<JObject> GetSiteDriveAsync(string siteName)
        {
            _logger.LogInformation("GetChannelDriveAsync called.");

            Guard.Against.NullOrEmpty(siteName, nameof(siteName));
            string result = string.Concat(siteName.Where(c => !char.IsWhiteSpace(c)));

            var siteIdResponse = await _graphSharePointAppService.GetSiteIdAsync(_appOptions.SharePointHostName, result);

            // Response field id is composed as follows: {hostname},{spsite.id},{spweb.id}
            var siteId = siteIdResponse["id"].ToString();

            var driveResponse = await _graphSharePointAppService.GetSiteDriveAsync(siteId);

            return driveResponse;
        }

        public async Task<JObject> GetSiteIdAsync(string siteName)
        {
            _logger.LogInformation("GetSiteIdAsync called.");

            Guard.Against.NullOrEmpty(siteName, nameof(siteName));
            string result = string.Concat(siteName.Where(c => !char.IsWhiteSpace(c)));

            var siteIdResponse = await _graphSharePointAppService.GetSiteIdAsync(_appOptions.SharePointHostName, result);

            return siteIdResponse;
        }
    }
}
