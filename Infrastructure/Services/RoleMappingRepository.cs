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
using ApplicationCore.Services;
using ApplicationCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Extensions.Caching.Memory;
using System.Linq;
using ApplicationCore.Helpers.Exceptions;

namespace Infrastructure.Services
{
    public class RoleMappingRepository : BaseRepository<RoleMapping>, IRoleMappingRepository
    {
        private readonly GraphSharePointAppService _graphSharePointAppService;
        private IMemoryCache _cache;

        public RoleMappingRepository(
            ILogger<RoleMappingRepository> logger,
            IOptions<AppOptions> appOptions,
            GraphSharePointAppService graphSharePointAppService,
            IMemoryCache memoryCache) : base(logger, appOptions)
        {
            Guard.Against.Null(graphSharePointAppService, nameof(graphSharePointAppService));

            _graphSharePointAppService = graphSharePointAppService;
            _cache = memoryCache;
        }

        public async Task<StatusCodes> CreateItemAsync(RoleMapping entity, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - RoleMappingRepo_CreateItemAsync called.");

            try
            {
                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.RolesListId
                };

                // Create Json object for SharePoint create list item
                dynamic itemFieldsJson = new JObject();
                itemFieldsJson.Title = entity.Id;
                itemFieldsJson.AdGroupName = entity.AdGroupName;
                itemFieldsJson.RoleName = entity.RoleName;
                itemFieldsJson.AdGroupId = entity.AdGroupId;
                itemFieldsJson.ProcessStep = entity.ProcessStep;
                itemFieldsJson.Channel = entity.Channel;
                itemFieldsJson.ProcessType = entity.ProcessType;


                dynamic itemJson = new JObject();
                itemJson.fields = itemFieldsJson;

                var result = await _graphSharePointAppService.CreateListItemAsync(siteList, itemJson.ToString(), requestId);

                _logger.LogInformation($"RequestId: {requestId} - RoleMappingRepo_CreateItemAsync finished creating SharePoint list item.");

                return StatusCodes.Status201Created;

            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - RoleMappingRepo_CreateItemAsync error: {ex}");
                throw;
            }
        }

        public async Task<StatusCodes> UpdateItemAsync(RoleMapping entity, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - RoleMappingRepo_UpdateItemAsync called.");

            try
            {
                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.RolesListId
                };

                // Create Json object for SharePoint create list item
                dynamic itemJson = new JObject();
                itemJson.Title = entity.Id;
                itemJson.AdGroupName = entity.AdGroupName;
                itemJson.RoleName = entity.RoleName;
                itemJson.AdGroupId = entity.AdGroupId;
                itemJson.ProcessStep = entity.ProcessStep;
                itemJson.Channel = entity.Channel;
                itemJson.ProcessType = entity.ProcessType;

                var result = await _graphSharePointAppService.UpdateListItemAsync(siteList, entity.Id, itemJson.ToString(), requestId);

                _logger.LogInformation($"RequestId: {requestId} - RoleMappingRepo_UpdateItemAsync finished creating SharePoint list item.");

                return StatusCodes.Status200OK;

            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - RoleMappingRepo_UpdateItemAsync error: {ex}");
                throw;
            }
        }

        public async Task<StatusCodes> DeleteItemAsync(string id, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - RoleMappingRepo_DeleteItemAsync called.");

            try
            {
                Guard.Against.NullOrEmpty(id, "RoleMappingRepo_DeleteItemAsync id null or empty", requestId);

                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.RolesListId
                };

                var result = await _graphSharePointAppService.DeleteListItemAsync(siteList, id, requestId);

                _logger.LogInformation($"RequestId: {requestId} - RoleMappingRepo_DeleteItemAsync finished creating SharePoint list item.");

                return StatusCodes.Status204NoContent;

            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - RoleMappingRepo_DeleteItemAsync error: {ex}");
                throw;
            }
        }

        public async Task<IList<RoleMapping>> GetAllAsync(string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - RoleMappingRepo_GetAllAsync called.");

            try
            {
                return await CacheTryGetRoleMappingListAsync(requestId);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - RoleMappingRepo_GetAllAsync error: {ex}");
                throw;
            }
        }

        private async Task<IList<RoleMapping>> CacheTryGetRoleMappingListAsync(string requestId = "")
        {
            try
            {
                var roleMappingList = new List<RoleMapping>();

                if (_appOptions.UserProfileCacheExpiration == 0)
                {
                    roleMappingList = (await GetRoleMappingListAsync(requestId)).ToList();
                }
                else
                {
                    var isExist = _cache.TryGetValue("PM_RoleMappingList", out roleMappingList);

                    if (!isExist)
                    {
                        roleMappingList = (await GetRoleMappingListAsync(requestId)).ToList();

                        var cacheEntryOptions = new MemoryCacheEntryOptions()
                            .SetAbsoluteExpiration(TimeSpan.FromMinutes(_appOptions.UserProfileCacheExpiration));

                        _cache.Set("PM_RoleMappingList", roleMappingList, cacheEntryOptions);
                    }
                }

                return roleMappingList;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - RoleMappingRepo_CacheTryGetRoleMappingListAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - RoleMappingRepo_CacheTryGetRoleMappingListAsync Service Exception: {ex}");
            }
        }

        private async Task<IList<RoleMapping>> GetRoleMappingListAsync(string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - RoleMappingRepo_GetRoleMappingListAsync called.");

            try
            {
                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.RolesListId
                };

                var json = await _graphSharePointAppService.GetListItemsAsync(siteList, "all", requestId);
                JArray jsonArray = JArray.Parse(json["value"].ToString());

                var itemsList = new List<RoleMapping>();
                foreach (var item in jsonArray)
                {
                    itemsList.Add(JsonConvert.DeserializeObject<RoleMapping>(item["fields"].ToString(), new JsonSerializerSettings
                    {
                        MissingMemberHandling = MissingMemberHandling.Ignore
                    }));
                }

                return itemsList;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - RoleMappingRepo_GetRoleMappingListAsync error: {ex}");
                throw;
            }
        }
    }
}
