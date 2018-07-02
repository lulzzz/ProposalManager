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
using Infrastructure.Services;
using ApplicationCore.Helpers;
using ApplicationCore.Entities.GraphServices;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ApplicationCore.Services
{
    public class IndustryRepository : BaseRepository<Industry>, IIndustryRepository
    {
        private readonly GraphSharePointAppService _graphSharePointAppService;

        public IndustryRepository(
            ILogger<IndustryRepository> logger, 
            IOptions<AppOptions> appOptions,
            GraphSharePointAppService graphSharePointAppService) : base(logger, appOptions)
        {
            Guard.Against.Null(graphSharePointAppService, nameof(graphSharePointAppService));
            _graphSharePointAppService = graphSharePointAppService;
        }

        public async Task<StatusCodes> CreateItemAsync(Industry entity, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - IndustryRepo_CreateItemAsync called.");

            try
            {
                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.IndustryListId
                };

                // Create Json object for SharePoint create list item
                dynamic itemFieldsJson = new JObject();
				itemFieldsJson.Title = entity.Id;
				itemFieldsJson.Name = entity.Name;

                dynamic itemJson = new JObject();
                itemJson.fields = itemFieldsJson;

                var result = await _graphSharePointAppService.CreateListItemAsync(siteList, itemJson.ToString(), requestId);

                _logger.LogInformation($"RequestId: {requestId} - IndustryRepo_CreateItemAsync finished creating SharePoint list item.");

                return StatusCodes.Status201Created;

            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - IndustryRepo_CreateItemAsync error: {ex}");
                throw;
            }
        }

        public async Task<StatusCodes> UpdateItemAsync(Industry entity, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - IndustryRepo_UpdateItemAsync called.");

            try
            {
                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.IndustryListId
                };

                // Create Json object for SharePoint create list item
                dynamic itemJson = new JObject();
                itemJson.Title = entity.Id;
                itemJson.Name = entity.Name;

                var result = await _graphSharePointAppService.UpdateListItemAsync(siteList, entity.Id, itemJson.ToString(), requestId);

                _logger.LogInformation($"RequestId: {requestId} - IndustryRepo_UpdateItemAsync finished creating SharePoint list item.");

                return StatusCodes.Status200OK;

            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - IndustryRepo_UpdateItemAsync error: {ex}");
                throw;
            }
        }

        public async Task<StatusCodes> DeleteItemAsync(string id, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - IndustryRepo_DeleteItemAsync called.");

            try
            {
                Guard.Against.NullOrEmpty(id, "IndustryRepo_DeleteItemAsync id null or empty", requestId);

                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.IndustryListId
                };

                var result = await _graphSharePointAppService.DeleteListItemAsync(siteList, id, requestId);

                _logger.LogInformation($"RequestId: {requestId} - IndustryRepo_DeleteItemAsync finished creating SharePoint list item.");

                return StatusCodes.Status204NoContent;

            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - IndustryRepo_DeleteItemAsync error: {ex}");
                throw;
            }
        }

        public async Task<IList<Industry>> GetAllAsync(string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - IndustryRepo_GetAllAsync called.");

            try
            {
                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.IndustryListId
                };

                var json = await _graphSharePointAppService.GetListItemsAsync(siteList, "all", requestId);
                JArray jsonArray = JArray.Parse(json["value"].ToString());

                var itemsList = new List<Industry>();
                foreach (var item in jsonArray)
                {
                    itemsList.Add(JsonConvert.DeserializeObject<Industry>(item["fields"].ToString(), new JsonSerializerSettings
                    {
                        MissingMemberHandling = MissingMemberHandling.Ignore
                    }));
                }

                return itemsList;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - IndustryRepo_GetAllAsync error: {ex}");
                throw;
            }
        }
    }
}
