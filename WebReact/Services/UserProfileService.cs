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
using ApplicationCore.Entities;
using Infrastructure.Services;
using ApplicationCore.Interfaces;
using ApplicationCore.Helpers;
using ApplicationCore.Helpers.Exceptions;
using WebReact.Models;
using WebReact.Helpers;

namespace WebReact.Services
{
	public class UserProfileService : BaseService<UserProfileService>, IUserProfileService
	{
		private readonly IUserProfileRepository _userProfileRepository;
        private readonly UserProfileHelpers _userProfileHelpers;

        public UserProfileService(
			ILogger<UserProfileService> logger,
			IOptions<AppOptions> appOptions,
			IUserProfileRepository userProfileRepository,
            UserProfileHelpers userProfileHelpers) : base(logger, appOptions)
		{
			Guard.Against.Null(userProfileRepository, nameof(userProfileRepository));
            Guard.Against.Null(userProfileHelpers, nameof(userProfileHelpers));

            _userProfileRepository = userProfileRepository;
            _userProfileHelpers = userProfileHelpers;

        }

		public async Task<UserProfileViewModel> GetItemByIdAsync(string id, string requestId = "")
		{
			_logger.LogInformation("GetItemByIdAsync called.");

			var selectedUserProfile = await _userProfileRepository.GetItemByIdAsync(id);
			var userProfileViewModel = await _userProfileHelpers.ToViewModelAsync(selectedUserProfile);

			return userProfileViewModel;
		}

		public async Task<UserProfileViewModel> GetItemByUpnAsync(string upn, string requestId = "")
		{
			_logger.LogInformation($"RequestId: {requestId} - GetItemByUpnAsync called.");

			try
			{
				var selectedUserProfile = await _userProfileRepository.GetItemByUpnAsync(upn, requestId);
				var userProfileViewModel = await _userProfileHelpers.ToViewModelAsync(selectedUserProfile);

				return userProfileViewModel;
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - GetItemByUpnAsync Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - GetItemByUpnAsync Service Exception: {ex}");
			}

		}

		public async Task<UserProfileListViewModel> GetAllAsync(int pageIndex, int itemsPage, string requestId = "")
		{
			_logger.LogInformation($"RequestId: {requestId} - GetAllAsync called.");

			try
			{
				var listItems = (await _userProfileRepository.GetAllAsync(requestId)).ToList();
				Guard.Against.Null(listItems, nameof(listItems), requestId);

				var userProfileListViewModel = new UserProfileListViewModel();
				foreach (var item in listItems)
				{
					userProfileListViewModel.ItemsList.Add(await _userProfileHelpers.ToViewModelAsync(item));
				}

				if (userProfileListViewModel.ItemsList.Count == 0)
				{
					_logger.LogWarning($"RequestId: {requestId} - GetAllAsync no items found");
					throw new NoItemsFound($"RequestId: {requestId} - Method name: GetAllAsync - No Items Found");
				}

				return userProfileListViewModel;
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - GetAllAsync Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - GetAllAsync Service Exception: {ex}");
			}
		}
	}
}
