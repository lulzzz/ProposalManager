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

namespace WebReact.Helpers
{
    public class UserProfileHelpers
    {
        protected readonly ILogger _logger;
        protected readonly AppOptions _appOptions;

        /// <summary>
        /// Constructor
        /// </summary>
        public UserProfileHelpers(
            ILogger<UserProfileHelpers> logger,
            IOptions<AppOptions> appOptions)
        {
            Guard.Against.Null(logger, nameof(logger));
            Guard.Against.Null(appOptions, nameof(appOptions));

            _logger = logger;
            _appOptions = appOptions.Value;
        }

        public async Task<UserProfile> UserProfileToEntityAsync(UserProfileViewModel model, string requestId = "")
        {
            var userProfile = UserProfile.Empty;

            userProfile.Id = model.Id ?? String.Empty;
            userProfile.DisplayName = model.DisplayName ?? String.Empty;
            userProfile.Fields = UserProfileFields.Empty;
            userProfile.Fields.Mail = model.Mail ?? String.Empty;
            userProfile.Fields.UserPrincipalName = model.UserPrincipalName ?? String.Empty;
            userProfile.Fields.Title = model.Title ?? String.Empty;
            userProfile.Fields.UserRoles = await RolesModelToEntitiesAsync(model.UserRoles, requestId);

            return userProfile;
        }

        public async Task<List<Role>> RolesModelToEntitiesAsync(List<RoleModel> roles, string requestId = "")
        {
            try
            {
                var rolesDto = new List<Role>();

                foreach (var itm in roles)
                {
                    var role = await RoleModelToEntityAsync(itm, requestId);

                    rolesDto.Add(role);
                }

                return rolesDto;
            }
            catch (Exception ex)
            {
                // TODO: _logger.LogError("MapToViewModelAsync error: " + ex);
                throw new ResponseException($"RequestId: {requestId} - RolesModelToEntitiesAsync - failed to map opportunity: {ex}");
            }
        }

        public Task<Role> RoleModelToEntityAsync(RoleModel role, string requestId = "")
        {
            try
            {
                var roleEntity = Role.Empty;
                roleEntity.Id = role.Id;
                roleEntity.DisplayName = role.DisplayName;
                roleEntity.AdGroupName = role.AdGroupName;

                return Task.FromResult(roleEntity);
            }
            catch (Exception ex)
            {
                // TODO: _logger.LogError("MapToViewModelAsync error: " + ex);
                throw new ResponseException($"RequestId: {requestId} - RoleModelToEntityAsync - failed to map opportunity: {ex}");
            }
        }

        public Task<UserProfileViewModel> ToViewModelAsync(UserProfile userProfile, string requestId = "")
        {
            try
            {
                var userProfileViewModel = new UserProfileViewModel
                {
                    Id = userProfile.Id,
                    DisplayName = userProfile.DisplayName,
                    Mail = userProfile.Fields.Mail ?? String.Empty,
                    UserPrincipalName = userProfile.Fields.UserPrincipalName,
                    Title = userProfile.Fields.Title ?? String.Empty,
                    UserRoles = new List<RoleModel>()
                };

                if (userProfile.Fields.UserRoles != null)
                {
                    foreach (var role in userProfile.Fields.UserRoles)
                    {
                        var userRole = new RoleModel();
                        userRole.Id = role.Id;
                        userRole.DisplayName = role.DisplayName;
                        userRole.AdGroupName = role.AdGroupName;

                        userProfileViewModel.UserRoles.Add(userRole);
                    }
                }

                return Task.FromResult(userProfileViewModel);
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - ToViewModel Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - ToViewModel Service Exception: {ex}");
            }
        }

        public async Task<List<RoleModel>> RolesToViewModelAsync(List<Role> roles, string requestId = "")
        {
            try
            {
                var rolesModel = new List<RoleModel>();

                foreach (var itm in roles)
                {
                    var roleModel = await RoleToViewModelAsync(itm, requestId);

                    rolesModel.Add(roleModel);
                }

                return rolesModel;
            }
            catch (Exception ex)
            {
                // TODO: _logger.LogError("MapToViewModelAsync error: " + ex);
                throw new ResponseException($"RequestId: {requestId} - RoleToViewModelAsync - failed to map opportunity: {ex}");
            }
        }

        public Task<RoleModel> RoleToViewModelAsync(Role role, string requestId = "")
        {
            try
            {
                var roleModel = new RoleModel();
                roleModel.Id = role.Id;
                roleModel.DisplayName = role.DisplayName;
                roleModel.AdGroupName = role.AdGroupName;

                return Task.FromResult(roleModel);
            }
            catch (Exception ex)
            {
                // TODO: _logger.LogError("MapToViewModelAsync error: " + ex);
                throw new ResponseException($"RequestId: {requestId} - RoleToViewModelAsync - failed to map opportunity: {ex}");
            }
        }
    }
}
