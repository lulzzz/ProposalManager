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
    public class CategoryService : BaseService<CategoryService>, ICategoryService
    {
        private readonly ICategoryRepository _categoryRepository;

        public CategoryService(
            ILogger<CategoryService> logger,
            IOptions<AppOptions> appOptions,
            ICategoryRepository categoryRepository) : base(logger, appOptions)
        {
            Guard.Against.Null(categoryRepository, nameof(categoryRepository));
            _categoryRepository = categoryRepository;
        }

        public async Task<StatusCodes> CreateItemAsync(CategoryModel modelObject, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - Category_CreateItemAsync called.");

            Guard.Against.Null(modelObject, nameof(modelObject), requestId);
            Guard.Against.NullOrEmpty(modelObject.Name, nameof(modelObject.Name), requestId);
            try
            {
                var entityObject = MapToEntity(modelObject, requestId);

                var result = await _categoryRepository.CreateItemAsync(entityObject, requestId);

                Guard.Against.NotStatus201Created(result, "Category_CreateItemAsync", requestId);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - Category_CreateItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - Category_CreateItemAsync Service Exception: {ex}");
            }
        }

        public async Task<StatusCodes> UpdateItemAsync(CategoryModel modelObject, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - Category_UpdateItemAsync called.");

            Guard.Against.Null(modelObject, nameof(modelObject), requestId);
            Guard.Against.NullOrEmpty(modelObject.Id, nameof(modelObject.Id), requestId);

            try
            {
                var entityObject = MapToEntity(modelObject, requestId);

                var result = await _categoryRepository.UpdateItemAsync(entityObject, requestId);

                Guard.Against.NotStatus200OK(result, "Category_UpdateItemAsync", requestId);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - Category_UpdateItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - Category_UpdateItemAsync Service Exception: {ex}");
            }
        }

        public async Task<StatusCodes> DeleteItemAsync(string id, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - Category_DeleteItemAsync called.");
            Guard.Against.NullOrEmpty(id, nameof(id), requestId);

            try
            {
                var result = await _categoryRepository.DeleteItemAsync(id, requestId);

                Guard.Against.NotStatus204NoContent(result, $"Category_DeleteItemAsync failed for id: {id}", requestId);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - Category_DeleteItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - Category_DeleteItemAsync Service Exception: {ex}");
            }
        }

        public async Task<IList<CategoryModel>> GetAllAsync(string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - CategorySvc_GetAllAsync called.");

            try
            {
                var listItems = (await _categoryRepository.GetAllAsync(requestId)).ToList();
                Guard.Against.Null(listItems, nameof(listItems), requestId);

                var modelListItems = new List<CategoryModel>();
                foreach (var item in listItems)
                {
                    modelListItems.Add(MapToModel(item));
                }

                if (modelListItems.Count == 0)
                {
                    _logger.LogWarning($"RequestId: {requestId} - CategorySvc_GetAllAsync no items found");
                    throw new NoItemsFound($"RequestId: {requestId} - Method name: CategorySvc_GetAllAsync - No Items Found");
                }

                return modelListItems;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - CategorySvc_GetAllAsync error: " + ex);
                throw;
            }
        }

        private CategoryModel MapToModel(Category entity, string requestId = "")
        {
            // Perform mapping
            var model = new CategoryModel();

            model.Id = entity.Id ?? String.Empty;
            model.Name = entity.Name ?? String.Empty;

            return model;
        }

        private Category MapToEntity(CategoryModel model, string requestId = "")
        {
            // Perform mapping
            var entity = Category.Empty;

            entity.Id = model.Id ?? String.Empty;
            entity.Name = model.Name ?? String.Empty;

            return entity;
        }
    }
}
