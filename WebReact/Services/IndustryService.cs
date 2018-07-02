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
    public class IndustryService : BaseService<IndustryService>, IIndustryService
    {
        private readonly IIndustryRepository _industryRepository;

        public IndustryService(
            ILogger<IndustryService> logger, 
            IOptions<AppOptions> appOptions,
            IIndustryRepository industryRepository) : base(logger, appOptions)
        {
            Guard.Against.Null(industryRepository, nameof(industryRepository));
            _industryRepository = industryRepository;
        }

        public async Task<StatusCodes> CreateItemAsync(IndustryModel modelObject, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - Industry_CreateItemAsync called.");

            Guard.Against.Null(modelObject, nameof(modelObject), requestId);
            Guard.Against.NullOrEmpty(modelObject.Name, nameof(modelObject.Name), requestId);
            try
            {
                var entityObject = MapToEntity(modelObject, requestId);

                var result = await _industryRepository.CreateItemAsync(entityObject, requestId);

                Guard.Against.NotStatus201Created(result, "Industry_CreateItemAsync", requestId);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - Industry_CreateItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - Industry_CreateItemAsync Service Exception: {ex}");
            }
        }

        public async Task<StatusCodes> UpdateItemAsync(IndustryModel modelObject, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - Industry_UpdateItemAsync called.");

            Guard.Against.Null(modelObject, nameof(modelObject), requestId);
            Guard.Against.NullOrEmpty(modelObject.Id, nameof(modelObject.Id), requestId);

            try
            {
                var entityObject = MapToEntity(modelObject, requestId);

                var result = await _industryRepository.UpdateItemAsync(entityObject, requestId);

                Guard.Against.NotStatus200OK(result, "Industry_UpdateItemAsync", requestId);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - Industry_UpdateItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - Industry_UpdateItemAsync Service Exception: {ex}");
            }
        }

        public async Task<StatusCodes> DeleteItemAsync(string id, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - Industry_DeleteItemAsync called.");
            Guard.Against.NullOrEmpty(id, nameof(id), requestId);

            try
            {
                var result = await _industryRepository.DeleteItemAsync(id, requestId);

                Guard.Against.NotStatus204NoContent(result, $"Industry_DeleteItemAsync failed for id: {id}", requestId);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - Industry_DeleteItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - Industry_DeleteItemAsync Service Exception: {ex}");
            }
        }

        public async Task<IList<IndustryModel>> GetAllAsync(string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - GetAllAsync called.");

            try
            {
                var listItems = (await _industryRepository.GetAllAsync(requestId)).ToList();
                Guard.Against.Null(listItems, nameof(listItems), requestId);

                var modelListItems = new List<IndustryModel>();
                foreach (var item in listItems)
                {
                    modelListItems.Add(MapToModel(item));
                }

                if (modelListItems.Count == 0)
                {
                    _logger.LogWarning($"RequestId: {requestId} - GetAllAsync no items found");
                    throw new NoItemsFound($"RequestId: {requestId} - Method name: GetAllAsync - No Items Found");
                }

                return modelListItems;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - GetAllAsync error: " + ex);
                throw;
            }
        }

        private IndustryModel MapToModel(Industry entity, string requestId = "")
        {
            // Perform mapping
            var model = new IndustryModel();

            model.Id = entity.Id ?? String.Empty;
            model.Name = entity.Name ?? String.Empty;

            return model;
        }

        private Industry MapToEntity(IndustryModel model, string requestId = "")
        {
            // Perform mapping
            var entity = Industry.Empty;

            entity.Id = model.Id ?? String.Empty;
            entity.Name = model.Name ?? String.Empty;

            return entity;
        }
    }
}
