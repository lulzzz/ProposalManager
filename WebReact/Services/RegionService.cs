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
using ApplicationCore.Helpers.Exceptions;

namespace WebReact.Services
{
    public class RegionService : BaseService<RegionService>, IRegionService
    {
        private readonly IRegionRepository _regionRepository;

        public RegionService(
            ILogger<RegionService> logger, 
            IOptions<AppOptions> appOptions,
            IRegionRepository regionRepository) : base(logger, appOptions)
        {
            Guard.Against.Null(regionRepository, nameof(regionRepository));
            _regionRepository = regionRepository;
        }


        public async Task<StatusCodes> CreateItemAsync(RegionModel modelObject, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - Region_CreateItemAsync called.");

            Guard.Against.Null(modelObject, nameof(modelObject), requestId);
            Guard.Against.NullOrEmpty(modelObject.Name, nameof(modelObject.Name), requestId);
            try
            {
                var entityObject = MapToEntity(modelObject, requestId);

                var result = await _regionRepository.CreateItemAsync(entityObject, requestId);

                Guard.Against.NotStatus201Created(result, "Region_CreateItemAsync", requestId);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - Region_CreateItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - Region_CreateItemAsync Service Exception: {ex}");
            }
        }

        public async Task<StatusCodes> UpdateItemAsync(RegionModel modelObject, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - Region_UpdateItemAsync called.");

            Guard.Against.Null(modelObject, nameof(modelObject), requestId);
            Guard.Against.NullOrEmpty(modelObject.Id, nameof(modelObject.Id), requestId);

            try
            {
                var entityObject = MapToEntity(modelObject, requestId);

                var result = await _regionRepository.UpdateItemAsync(entityObject, requestId);

                Guard.Against.NotStatus200OK(result, "Region_UpdateItemAsync", requestId);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - Region_UpdateItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - Region_UpdateItemAsync Service Exception: {ex}");
            }
        }

        public async Task<StatusCodes> DeleteItemAsync(string id, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - Region_DeleteItemAsync called.");
            Guard.Against.NullOrEmpty(id, nameof(id), requestId);

            try
            {
                var result = await _regionRepository.DeleteItemAsync(id, requestId);

                Guard.Against.NotStatus204NoContent(result, $"Region_DeleteItemAsync failed for id: {id}", requestId);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - Region_DeleteItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - Region_DeleteItemAsync Service Exception: {ex}");
            }
        }

        public async Task<IList<RegionModel>> GetAllAsync(string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - RegionSvc_GetAllAsync called.");

            try
            {
                var listItems = (await _regionRepository.GetAllAsync(requestId)).ToList();
                Guard.Against.Null(listItems, nameof(listItems), requestId);

                var modelListItems = new List<RegionModel>();
                foreach (var item in listItems)
                {
                    modelListItems.Add(MapToModel(item));
                }

                if (modelListItems.Count == 0)
                {
                    _logger.LogWarning($"RequestId: {requestId} - RegionSvc_GetAllAsync no items found");
                    throw new NoItemsFound($"RequestId: {requestId} - Method name: RegionSvc_GetAllAsync - No Items Found");
                }

                return modelListItems;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - RegionSvc_GetAllAsync error: " + ex);
                throw;
            }
        }

        private RegionModel MapToModel(Region entity, string requestId = "")
        {
            // Perform mapping
            var model = new RegionModel();

            model.Id = entity.Id ?? String.Empty;
            model.Name = entity.Name ?? String.Empty;

            return model;
        }

        private Region MapToEntity(RegionModel model, string requestId = "")
        {
            // Perform mapping
            var entity = Region.Empty;

            entity.Id = model.Id ?? String.Empty;
            entity.Name = model.Name ?? String.Empty;

            return entity;
        }
    }
}
