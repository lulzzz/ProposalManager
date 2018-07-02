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
using ApplicationCore.Helpers.Exceptions;
using Microsoft.Extensions.Caching.Memory;

namespace Infrastructure.Services
{
	public class UserProfileRepository : BaseRepository<UserProfile>, IUserProfileRepository
	{
		private IMemoryCache _cache;
		private readonly GraphSharePointAppService _graphSharePointAppService;
		private readonly GraphUserAppService _graphUserAppService;
        private readonly IRoleMappingRepository _roleMappingRepository;
		private JArray _roleMappingList;
		private List<UserProfile> _usersList;

		public UserProfileRepository(ILogger<UserProfileRepository> logger,
			 GraphSharePointAppService graphSharePointAppService,
			 GraphUserAppService graphUserAppService,
             IRoleMappingRepository roleMappingRepository,
             IOptions<AppOptions> appOptions,
			 IMemoryCache memoryCache) : base(logger, appOptions)
		{
			Guard.Against.Null(graphSharePointAppService, nameof(graphSharePointAppService));
            Guard.Against.Null(graphUserAppService, nameof(graphUserAppService));
            Guard.Against.Null(roleMappingRepository, nameof(roleMappingRepository));

            _graphSharePointAppService = graphSharePointAppService;
			_graphUserAppService = graphUserAppService;
            _roleMappingRepository = roleMappingRepository;
            _cache = memoryCache;

			_roleMappingList = null;
			_usersList = new List<UserProfile>();
		}

		public async Task<UserProfile> GetItemByIdAsync(string id, string requestId = "")
		{
			_logger.LogInformation($"RequestId: {requestId} - GetItemByIdAsync called.");

			try
			{
				Guard.Against.NullOrEmpty(id, "GetItemByIdAsync_id Null", requestId);

				var usersList = await CacheTryGetUsersListAsync(requestId);

				var userProfile = usersList.Find(x => x.Id == id);
				if (userProfile == null)
				{
					_logger.LogWarning($"RequestId: {requestId} - GetItemByIdAsync_id no user found: {id}");
					throw new ResponseException($"RequestId: {requestId} - GetItemByIdAsync_id Sno user found: {id}");
				}

				return userProfile;
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - GetItemByIdAsync Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - GetItemByIdAsync Service Exception: {ex}");
			};
		}

		public async Task<UserProfile> GetItemByUpnAsync(string upn, string requestId = "")
		{
			_logger.LogInformation($"RequestId: {requestId} - GetItemByIdAsync called.");

			try
			{
				Guard.Against.NullOrEmpty(upn, "GetItemByUmpn_upn Null", requestId);

				var usersList = await CacheTryGetUsersListAsync(requestId);

				var userProfile = usersList.Find(x => x.Fields.UserPrincipalName == upn);
				if (userProfile == null)
				{
					_logger.LogWarning($"RequestId: {requestId} - GetItemByUpnAsync no user found: {upn}");
					throw new ResponseException($"RequestId: {requestId} - GetItemByUpnAsync Sno user found: {upn}");
				}

				return userProfile;
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - GetItemByUpnAsync Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - GetItemByUpnAsync Service Exception: {ex}");
			}

		}

		public async Task<IList<UserProfile>> GetAllAsync(string requestId = "")
		{
			_logger.LogInformation($"RequestId: {requestId} - GetAllAsync called.");

			try
			{
				return await CacheTryGetUsersListAsync(requestId);
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - GetAllAsync Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - GetAllAsync Service Exception: {ex}");
			}
		}


		// Private methods
		private async Task<JArray> GetRoleMappingListAsync(string requestId = "")
		{
			try
			{
				if (_roleMappingList == null)
				{
					var siteList = new SiteList
					{
						SiteId = _appOptions.ProposalManagementRootSiteId,
						ListId = _appOptions.RolesListId
					};

					var json = await _graphSharePointAppService.GetListItemsAsync(siteList, "all", requestId);
					JArray jArrayResult = JArray.Parse(json["value"].ToString());

					_roleMappingList = jArrayResult;
				}

				return _roleMappingList;
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - GetRoleMappingList Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - GetRoleMappingList Service Exception: {ex}");
			}
		}

		private async Task<List<UserProfile>> GetUsersListAsync(string requestId = "")
		{
			try
			{
				if (_usersList?.Count == 0)
				{
					var roleMappings = await _roleMappingRepository.GetAllAsync(requestId);
					foreach(var role in roleMappings)
					{
						var userRole = Role.Empty;
						userRole.DisplayName = role.RoleName;
						userRole.AdGroupName = role.AdGroupName;

						var options = new List<QueryParam>();
						options.Add(new QueryParam("filter", $"startswith(displayName,'{userRole.AdGroupName}')"));
						var groupIdJson = await _graphUserAppService.GetGroupAsync(options, "", requestId);
						dynamic jsonDyn = groupIdJson;
						if (jsonDyn.value.HasValues)
						{
							userRole.Id = jsonDyn.value[0].id.ToString();

							var groupMembersJson = await _graphUserAppService.GetGroupMembersAsync(userRole.Id, requestId);
							JArray membersJsonArray = JArray.Parse(groupMembersJson["value"].ToString());

							foreach (var member in membersJsonArray)
							{
								var user = UserProfile.Empty;
								user = _usersList.Find(x => x.Id == member["id"].ToString());

								if (user != null)
								{
									_usersList.Remove(user);
								}
								else
								{
									user = UserProfile.Empty;
									user.Id = member["id"].ToString();
								}

								user.DisplayName = member["displayName"].ToString();
								if (user.Fields == null) user.Fields = UserProfileFields.Empty;
								user.Fields.Mail = member["mail"].ToString();
								user.Fields.UserPrincipalName = member["userPrincipalName"].ToString();
                                user.Fields.Title = member["jobTitle"].ToString() ?? String.Empty;

                                // Check if user already has the role
                                var existingRole = user.Fields.UserRoles.Find(x => x.Id == userRole.Id);
                                if (existingRole == null)
                                {
                                    user.Fields.UserRoles.Add(userRole);
                                }

								_usersList.Add(user);
							}
						}
					}
				}

				return _usersList;
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - GetUsersListAsync Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - GetUsersListAsync Service Exception: {ex}");
			}
		}

		private async Task<List<UserProfile>> CacheTryGetUsersListAsync(string requestId = "")
		{
			try
			{
				var userProfileList = new List<UserProfile>();

				if (_appOptions.UserProfileCacheExpiration == 0)
				{
					userProfileList = await GetUsersListAsync(requestId);
				}
				else
				{
					var isExist = _cache.TryGetValue("PM_UsersList", out userProfileList);

					if (!isExist)
					{
						userProfileList = await GetUsersListAsync(requestId);

						var cacheEntryOptions = new MemoryCacheEntryOptions()
							.SetAbsoluteExpiration(TimeSpan.FromMinutes(_appOptions.UserProfileCacheExpiration));

						_cache.Set("PM_UsersList", userProfileList, cacheEntryOptions);
					}
				}

				return userProfileList;
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - CacheGetOrCreateUsersListAsync Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - CacheGetOrCreateUsersListAsync Service Exception: {ex}");
			}
		}
	}
}
