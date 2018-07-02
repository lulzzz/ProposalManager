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
using ApplicationCore.Services;
using ApplicationCore;
using ApplicationCore.Helpers;
using ApplicationCore.Entities.GraphServices;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;

namespace Infrastructure.Services
{
    public class CategoryRepository : BaseRepository<Category>, ICategoryRepository
    {
        private readonly GraphSharePointAppService _graphSharePointAppService;

        public CategoryRepository(
            ILogger<CategoryRepository> logger, 
            IOptions<AppOptions> appOptions,
            GraphSharePointAppService graphSharePointAppService) : base(logger, appOptions)
        {
            Guard.Against.Null(graphSharePointAppService, nameof(graphSharePointAppService));
            _graphSharePointAppService = graphSharePointAppService;
        }


        public async Task<StatusCodes> CreateItemAsync(Category entity, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CategoryRepo_CreateItemAsync called.");

            try
            {
                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.CategoriesListId
                };

                // Create Json object for SharePoint create list item
                dynamic itemFieldsJson = new JObject();
                itemFieldsJson.Name = entity.Name;
				itemFieldsJson.Title = entity.Id;

				dynamic itemJson = new JObject();
                itemJson.fields = itemFieldsJson;

                var result = await _graphSharePointAppService.CreateListItemAsync(siteList, itemJson.ToString(), requestId);

                _logger.LogInformation($"RequestId: {requestId} - CategoryRepo_CreateItemAsync finished creating SharePoint list item.");

                return StatusCodes.Status201Created;

            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - CategoryRepo_CreateItemAsync error: {ex}");
                throw;
            }
        }

        public async Task<StatusCodes> UpdateItemAsync(Category entity, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CategoryRepo_UpdateItemAsync called.");

            try
            {
                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.CategoriesListId
                };

                // Create Json object for SharePoint create list item
                dynamic itemJson = new JObject();
                itemJson.Title = entity.Id;
                itemJson.Name = entity.Name;

                var result = await _graphSharePointAppService.UpdateListItemAsync(siteList, entity.Id, itemJson.ToString(), requestId);

                _logger.LogInformation($"RequestId: {requestId} - CategoryRepo_UpdateItemAsync finished creating SharePoint list item.");

                return StatusCodes.Status200OK;

            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - CategoryRepo_UpdateItemAsync error: {ex}");
                throw;
            }
        }

        public async Task<StatusCodes> DeleteItemAsync(string id, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CategoryRepo_DeleteItemAsync called.");

            try
            {
                Guard.Against.NullOrEmpty(id, "CategoryRepo_DeleteItemAsync id null or empty", requestId);

                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.CategoriesListId
                };

                var result = await _graphSharePointAppService.DeleteListItemAsync(siteList, id, requestId);

                _logger.LogInformation($"RequestId: {requestId} - CategoryRepo_DeleteItemAsync finished creating SharePoint list item.");

                return StatusCodes.Status204NoContent;

            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - CategoryRepo_DeleteItemAsync error: {ex}");
                throw;
            }
        }

        public async Task<IList<Category>> GetAllAsync(string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CategoryRepo_GetAllAsync called.");

            try
            {
                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.CategoriesListId
                };

                var json = await _graphSharePointAppService.GetListItemsAsync(siteList, "all", requestId);
                JArray jsonArray = JArray.Parse(json["value"].ToString());

                var itemsList = new List<Category>();
                foreach (var item in jsonArray)
                {
                    itemsList.Add(JsonConvert.DeserializeObject<Category>(item["fields"].ToString(), new JsonSerializerSettings
                    {
                        MissingMemberHandling = MissingMemberHandling.Ignore
                    }));
                }

                return itemsList;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - CategoryRepo_GetAllAsync error: {ex}");
                throw;
            }
        }
    }
}
