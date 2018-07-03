// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ApplicationCore.Artifacts;
using ApplicationCore.Interfaces;
using ApplicationCore.Entities;
using ApplicationCore.Services;
using ApplicationCore;
using ApplicationCore.Helpers;
using ApplicationCore.Entities.GraphServices;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using Infrastructure.GraphApi;
using ApplicationCore.Helpers.Exceptions;
using Microsoft.Graph;
using System.Linq;
using System.Text.RegularExpressions;

namespace Infrastructure.Services
{
    public class OpportunityRepository : BaseArtifactFactory<Opportunity>, IOpportunityRepository
    {
        private readonly IOpportunityFactory _opportunityFactory;
        private readonly GraphSharePointAppService _graphSharePointAppService;
        private readonly GraphUserAppService _graphUserAppService;
        private readonly INotificationRepository _notificationRepository;
        private readonly IUserProfileRepository _userProfileRepository;
        private readonly IUserContext _userContext;


        public OpportunityRepository(
            ILogger<OpportunityRepository> logger,
            IOptions<AppOptions> appOptions,
            GraphSharePointAppService graphSharePointAppService,
            GraphUserAppService graphUserAppService,
            INotificationRepository notificationRepository,
            IUserProfileRepository userProfileRepository,
            IUserContext userContext,
            IOpportunityFactory opportunityFactory) : base(logger, appOptions)
        {
            Guard.Against.Null(graphSharePointAppService, nameof(graphSharePointAppService));
            Guard.Against.Null(graphUserAppService, nameof(graphUserAppService));
            Guard.Against.Null(notificationRepository, nameof(notificationRepository));
            Guard.Against.Null(userProfileRepository, nameof(userProfileRepository));
            Guard.Against.Null(userContext, nameof(userContext));
            Guard.Against.Null(opportunityFactory, nameof(opportunityFactory));

            _graphSharePointAppService = graphSharePointAppService;
            _graphUserAppService = graphUserAppService;
            _notificationRepository = notificationRepository;
            _userProfileRepository = userProfileRepository;
            _userContext = userContext;
            _opportunityFactory = opportunityFactory;
        }

        public async Task<StatusCodes> CreateItemAsync(Opportunity opportunity, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - OpportunityRepository_CreateItemAsync called.");

            try
            {
                Guard.Against.Null(opportunity, nameof(opportunity), requestId);
                Guard.Against.NullOrEmpty(opportunity.DisplayName, nameof(opportunity.DisplayName), requestId);

                // Check access
                var roles = new List<Role>();
                roles.Add(new Role { DisplayName = "RelationshipManager" });
                var checkAccess = await _opportunityFactory.CheckAccessAsync(opportunity, roles, requestId);
                if (!checkAccess) _logger.LogError($"RequestId: {requestId} - OpportunityRepository_CreateItemAsync CheckAccess CreateItemAsync");

                // Ensure id is blank since it will be set by SharePoint
                opportunity.Id = String.Empty;

                // TODO: This section will be replaced with a workflow
                opportunity = await _opportunityFactory.CreateWorkflowAsync(opportunity, requestId);


                _logger.LogInformation($"RequestId: {requestId} - OpportunityRepository_CreateItemAsync creating SharePoint List for opportunity.");

                //Get loan officer & relationship manager values
                var loanOfficerId = String.Empty;
                var relationshipManagerId = String.Empty;
                var loanOfficerUpn = String.Empty;
                var relationshipManagerUpn = String.Empty;
                foreach (var item in opportunity.Content.TeamMembers)
                {
                    if (item.AssignedRole.DisplayName == "LoanOfficer" && !String.IsNullOrEmpty(item.Id))
                    {
                        loanOfficerId = item.Id;
                        loanOfficerUpn = item.Fields.UserPrincipalName;
                    }
                    if (item.AssignedRole.DisplayName == "RelationshipManager" && !String.IsNullOrEmpty(item.Id))
                    {
                        relationshipManagerId = item.Id;
                        relationshipManagerUpn = item.Fields.UserPrincipalName;
                    }
                }


                // Create Json object for SharePoint create list item
                dynamic opportunityFieldsJson = new JObject();
                //opportunityFieldsJson.OpportunityId = opportunity.Id;
                opportunityFieldsJson.Name = opportunity.DisplayName;
                opportunityFieldsJson.OpportunityState = opportunity.Metadata.OpportunityState.Name;
                opportunityFieldsJson.OpportunityObject = JsonConvert.SerializeObject(opportunity, Formatting.Indented);
                opportunityFieldsJson.LoanOfficer = loanOfficerId;
                opportunityFieldsJson.RelationshipManager = relationshipManagerId;

                dynamic opportunityJson = new JObject();
                opportunityJson.fields = opportunityFieldsJson;

                var opportunitySiteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.OpportunitiesListId
                };

                var result = await _graphSharePointAppService.CreateListItemAsync(opportunitySiteList, opportunityJson.ToString(), requestId);

                _logger.LogInformation($"RequestId: {requestId} - OpportunityRepository_CreateItemAsync finished creating SharePoint List for opportunity.");
                // END TODO

                // Add entry in Opportunities public sub site
                var opportunitySubSiteList = new SiteList
                {
                    SiteId = _appOptions.OpportunitiesSubSiteId,
                    ListId = _appOptions.PublicOpportunitiesListId
                };


                // Create Json object for SharePoint create list item
                dynamic pubOpportunityFieldsJson = new JObject();
                pubOpportunityFieldsJson.Title = opportunity.DisplayName;
                pubOpportunityFieldsJson.LoanOfficer = loanOfficerUpn;
                pubOpportunityFieldsJson.RelationshipManager = relationshipManagerUpn;
                pubOpportunityFieldsJson.State = opportunity.Metadata.OpportunityState.Name;

                dynamic pubOpportunityJson = new JObject();
                pubOpportunityJson.fields = pubOpportunityFieldsJson;

                try
                {
                    var resultPub = await _graphSharePointAppService.CreateListItemAsync(opportunitySubSiteList, pubOpportunityJson.ToString(), requestId);
                }
                catch (Exception ex)
                {
                    // Dont brak the opportunity creation of entry can't be added to subsite (public)
                    _logger.LogError($"RequestId: {requestId} - OpportunityRepository_CreateItemAsync SubsiteOpportunity Service Exception: {ex}");
                }


                return StatusCodes.Status201Created;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - OpportunityRepository_CreateItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - OpportunityRepository_CreateItemAsync Service Exception: {ex}");
            }
        }

        public async Task<StatusCodes> UpdateItemAsync(Opportunity opportunity, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - OpportunityRepository_UpdateItemAsync called.");
            Guard.Against.Null(opportunity, nameof(opportunity), requestId);
            Guard.Against.NullOrEmpty(opportunity.Id, nameof(opportunity.Id), requestId);

            try
            {
                // TODO: This section will be replaced with a workflow
                _logger.LogInformation($"RequestId: {requestId} - OpportunityRepository_UpdateItemAsync SharePoint List for opportunity.");

                // Check access
                var checkAccess = await _opportunityFactory.CheckAccessAnyAsync(opportunity, requestId);
                if (!checkAccess) _logger.LogError($"RequestId: {requestId} - OpportunityRepository_UpdateItemAsync CheckAccessAny");

                // Workflow processor
                opportunity = await _opportunityFactory.UpdateWorkflowAsync(opportunity, requestId);

                //Get loan officer & relationship manager values
                var loanOfficerId = String.Empty;
                var relationshipManagerId = String.Empty;
                var loanOfficerUpn = String.Empty;
                var relationshipManagerUpn = String.Empty;
                foreach (var item in opportunity.Content.TeamMembers)
                {
                    if (item.AssignedRole.DisplayName == "LoanOfficer" && !String.IsNullOrEmpty(item.Id))
                    {
                        loanOfficerId = item.Id;
                        loanOfficerUpn = item.Fields.UserPrincipalName;
                    }
                    if (item.AssignedRole.DisplayName == "RelationshipManager" && !String.IsNullOrEmpty(item.Id))
                    {
                        relationshipManagerId = item.Id;
                        relationshipManagerUpn = item.Fields.UserPrincipalName;
                    }
                }


                var opportunityJObject = JObject.FromObject(opportunity);

                // Create Json object for SharePoint create list item
                dynamic opportunityJson = new JObject();
                opportunityJson.OpportunityId = opportunity.Id;
                //opportunityJson.Name = opportunity.DisplayName; TODO: In wave 1 nme can't be changed
                opportunityJson.OpportunityState = opportunity.Metadata.OpportunityState.Name;
                opportunityJson.OpportunityObject = JsonConvert.SerializeObject(opportunity, Formatting.Indented);
                //opportunityJson.OpportunityObject = opportunityJObject.ToString();
                opportunityJson.LoanOfficer = loanOfficerId;
                opportunityJson.RelationshipManager = relationshipManagerId;

                var opportunitySiteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.OpportunitiesListId
                };
                var result = await _graphSharePointAppService.UpdateListItemAsync(opportunitySiteList, opportunity.Id, opportunityJson.ToString(), requestId);

                _logger.LogInformation($"RequestId: {requestId} - OpportunityRepository_UpdateItemAsync finished SharePoint List for opportunity.");


                // Update entry in Opportunities public sub site
                var opportunitySubSiteList = new SiteList
                {
                    SiteId = _appOptions.OpportunitiesSubSiteId,
                    ListId = _appOptions.PublicOpportunitiesListId
                };


                // Create Json object for SharePoint create list item
                dynamic pubOpportunityFieldsJson = new JObject();
                pubOpportunityFieldsJson.Title = opportunity.DisplayName;
                pubOpportunityFieldsJson.LoanOfficer = loanOfficerUpn;
                pubOpportunityFieldsJson.RelationshipManager = relationshipManagerUpn;
                pubOpportunityFieldsJson.State = opportunity.Metadata.OpportunityState.Name;

                var options = new List<QueryParam>();
                options.Add(new QueryParam("filter", $"startswith(fields/Title,'{opportunity.DisplayName}')"));

                try
                {
                    var json = await _graphSharePointAppService.GetListItemAsync(opportunitySubSiteList, options, "All", requestId);
                    Guard.Against.Null(json, "OpportunityRepository_UpdateItemAsync GetListItemAsync Null", requestId);
                    dynamic jsonDyn = json;
                    string opportunityId = jsonDyn.value[0].fields.id.ToString();
                    if (!String.IsNullOrEmpty(opportunityId))
                    {
                        var resultPub = await _graphSharePointAppService.UpdateListItemAsync(opportunitySubSiteList, opportunityId, pubOpportunityFieldsJson.ToString(), requestId);
                    }
                }
                catch (Exception ex)
                {
                    // Dont brak the opportunity creation of entry can't be updated to subsite (public)
                    _logger.LogError($"RequestId: {requestId} - OpportunityRepository_UpdateItemAsync SubsiteOpportunity Service Exception: {ex}");
                }

                return StatusCodes.Status200OK;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - OpportunityRepository_UpdateItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - OpportunityRepository_UpdateItemAsync Service Exception: {ex}");
            }
        }

        public async Task<Opportunity> GetItemByIdAsync(string id, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - OpportunityRepository_GetItemByIdAsync called.");

            try
            {
                Guard.Against.NullOrEmpty(id, nameof(id), requestId);

                var opportunitySiteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.OpportunitiesListId
                };

                var json = await _graphSharePointAppService.GetListItemByIdAsync(opportunitySiteList, id, "all", requestId);
                Guard.Against.Null(json, nameof(json), requestId);

                var opportunityJson = json["fields"]["OpportunityObject"].ToString();

                var oppArtifact = JsonConvert.DeserializeObject<Opportunity>(opportunityJson.ToString(), new JsonSerializerSettings
                {
                    MissingMemberHandling = MissingMemberHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore
                });

                oppArtifact.Id = json["fields"]["id"].ToString();

                // Check access
                var checkAccess = await _opportunityFactory.CheckAccessAnyAsync(oppArtifact, requestId);
                if (!checkAccess) _logger.LogError($"RequestId: {requestId} - OpportunityRepository_GetItemByIdAsync CheckAccessAny");

                return oppArtifact;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - OpportunityRepository_GetItemByIdAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - OpportunityRepository_GetItemByIdAsync Service Exception: {ex}");
            }
        }

        public async Task<Opportunity> GetItemByNameAsync(string name, bool isCheckName, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - OpportunityRepository_GetItemByNameAsync called.");

            try
            {
                Guard.Against.NullOrEmpty(name, nameof(name), requestId);

                var opportunitySiteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.OpportunitiesListId
                };

                name = name.Replace("'", "");
                var nameEncoded = WebUtility.UrlEncode(name);
                var options = new List<QueryParam>();
                options.Add(new QueryParam("filter", $"startswith(fields/Name,'{nameEncoded}')"));

                var json = await _graphSharePointAppService.GetListItemAsync(opportunitySiteList, options, "all", requestId);
                Guard.Against.Null(json, "OpportunityRepository_GetItemByNameAsync GetListItemAsync Null", requestId);

                dynamic jsonDyn = json;

                if (jsonDyn.value.HasValues)
                {
                    foreach (var item in jsonDyn.value)
                    {
                        if (item.fields.Name == name)
                        {
                            if (isCheckName)
                            {
                                // If just checking for name, rtunr empty opportunity and skip access check
                                var emptyOpportunity = Opportunity.Empty;
                                emptyOpportunity.DisplayName = name;
                                return emptyOpportunity;
                            }

                            var opportunityJson = item.fields.OpportunityObject.ToString();

                            var oppArtifact = JsonConvert.DeserializeObject<Opportunity>(opportunityJson, new JsonSerializerSettings
                            {
                                MissingMemberHandling = MissingMemberHandling.Ignore,
                                NullValueHandling = NullValueHandling.Ignore
                            });

                            oppArtifact.Id = jsonDyn.value[0].fields.id.ToString();

                            // Check access
                            var checkAccess = await _opportunityFactory.CheckAccessAnyAsync(oppArtifact, requestId);
                            if (!checkAccess) _logger.LogError($"RequestId: {requestId} - OpportunityRepository_GetItemByNameAsync CheckAccessAny");

                            return oppArtifact;
                        }
                    }

                }

                // Not found
                _logger.LogError($"RequestId: {requestId} - OpportunityRepository_GetItemByNameAsync opportunity: {name} - Not found.");

                return Opportunity.Empty;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - OpportunityRepository_GetItemByNameAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - OpportunityRepository_GetItemByNameAsync Service Exception: {ex}");
            }
        }

        public async Task<IList<Opportunity>> GetAllAsync(string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - OpportunityRepository_GetAllAsync called.");

            try
            {
                var siteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.OpportunitiesListId
                };

                var currentUser = (_userContext.User.Claims).ToList().Find(x => x.Type == "preferred_username")?.Value;
                Guard.Against.NullOrEmpty(currentUser, "OpportunityRepository_GetAllAsync CurrentUser null-empty", requestId);

                var callerUser = await _userProfileRepository.GetItemByUpnAsync(currentUser, requestId);
                Guard.Against.Null(callerUser, "_userProfileRepository.GetItemByUpnAsync Null", requestId);
                if (currentUser != callerUser.Fields.UserPrincipalName)
                {
                    _logger.LogError($"RequestId: {requestId} - OpportunityRepository_GetItemByIdAsync current user: {currentUser} AccessDeniedException");
                    throw new AccessDeniedException($"RequestId: {requestId} - OpportunityRepository_GetItemByIdAsync current user: {currentUser} AccessDeniedException");
                }

                var isLoanOfficer = false;
                var isRelationshipManager = false;
                var isAdmin = false;

                if (callerUser.Fields.UserRoles.Find(x => x.DisplayName == "LoanOfficer") != null)
                {
                    isLoanOfficer = true;
                }
                if (callerUser.Fields.UserRoles.Find(x => x.DisplayName == "RelationshipManager") != null)
                {
                    isRelationshipManager = true;
                }
                if (callerUser.Fields.UserRoles.Find(x => x.DisplayName == "Administrator") != null)
                {
                    isAdmin = true;
                }

                if (isLoanOfficer == false && isRelationshipManager == false && isAdmin == false)
                {
                    // This user is not LoannOfficer or RelationshipManager so it does not has access to list opportunities
                    _logger.LogError($"RequestId: {requestId} - OpportunityRepository_GetItemByIdAsync current user: {currentUser} AccessDeniedException");
                    throw new AccessDeniedException($"RequestId: {requestId} - OpportunityRepository_GetItemByIdAsync current user: {currentUser} AccessDeniedException");
                }

                var options = new List<QueryParam>();
                var jsonLoanOfficer = new JObject();
                var jsonRelationshipManager = new JObject();
                var jsonAdmin = new JObject();
                var itemsList = new List<Opportunity>();
                var jsonArray = new JArray();

                if (isAdmin)
                {
                    jsonAdmin = await _graphSharePointAppService.GetListItemsAsync(siteList, "all", requestId);

                    if (jsonAdmin.HasValues)
                    {
                        jsonArray = JArray.Parse(jsonAdmin["value"].ToString());
                    }


                    foreach (var item in jsonArray)
                    {
                        var opportunityJson = item["fields"]["OpportunityObject"].ToString();

                        var oppArtifact = JsonConvert.DeserializeObject<Opportunity>(opportunityJson.ToString(), new JsonSerializerSettings
                        {
                            MissingMemberHandling = MissingMemberHandling.Ignore,
                            NullValueHandling = NullValueHandling.Ignore
                        });

                        oppArtifact.Id = item["fields"]["id"].ToString();

                        itemsList.Add(oppArtifact);
                    }
                }
                else
                {
                    if (isLoanOfficer)
                    {
                        options.Add(new QueryParam("filter", $"startswith(fields/LoanOfficer,'{callerUser.Id}')"));
                        jsonLoanOfficer = await _graphSharePointAppService.GetListItemsAsync(siteList, options, "all", requestId);
                    }


                    if (isRelationshipManager)
                    {
                        options.Add(new QueryParam("filter", $"startswith(fields/RelationshipManager,'{callerUser.Id}')"));
                        jsonRelationshipManager = await _graphSharePointAppService.GetListItemsAsync(siteList, options, "all", requestId);
                    }

                    if (jsonLoanOfficer.HasValues)
                    {
                        jsonArray = JArray.Parse(jsonLoanOfficer["value"].ToString());
                    }


                    foreach (var item in jsonArray)
                    {
                        var opportunityJson = item["fields"]["OpportunityObject"].ToString();

                        var oppArtifact = JsonConvert.DeserializeObject<Opportunity>(opportunityJson.ToString(), new JsonSerializerSettings
                        {
                            MissingMemberHandling = MissingMemberHandling.Ignore,
                            NullValueHandling = NullValueHandling.Ignore
                        });

                        oppArtifact.Id = item["fields"]["id"].ToString();

                        itemsList.Add(oppArtifact);
                    }

                    if (jsonRelationshipManager.HasValues)
                    {
                        jsonArray = JArray.Parse(jsonRelationshipManager["value"].ToString());
                    }

                    foreach (var item in jsonArray)
                    {
                        var opportunityJson = item["fields"]["OpportunityObject"].ToString();

                        var oppArtifact = JsonConvert.DeserializeObject<Opportunity>(opportunityJson.ToString(), new JsonSerializerSettings
                        {
                            MissingMemberHandling = MissingMemberHandling.Ignore,
                            NullValueHandling = NullValueHandling.Ignore
                        });

                        oppArtifact.Id = item["fields"]["id"].ToString();

                        var dupeOpp = itemsList.Find(x => x.DisplayName == oppArtifact.DisplayName);
                        if (dupeOpp == null)
                        {
                            itemsList.Add(oppArtifact);
                        }
                    }
                }

                return itemsList;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - OpportunityRepository_GetAllAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - OpportunityRepository_GetAllAsync Service Exception: {ex}");
            }
        }

        public async Task<StatusCodes> DeleteItemAsync(string id, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - OpportunityRepository_DeleteItemAsync called.");

            try
            {
                Guard.Against.Null(id, nameof(id), requestId);

                var opportunitySiteList = new SiteList
                {
                    SiteId = _appOptions.ProposalManagementRootSiteId,
                    ListId = _appOptions.OpportunitiesListId
                };

                var opportunity = await _graphSharePointAppService.GetListItemByIdAsync(opportunitySiteList, id, "all", requestId);
                Guard.Against.Null(opportunity, $"OpportunityRepository_y_DeleteItemsAsync getItemsById: {id}", requestId);

                var opportunityJson = opportunity["fields"]["OpportunityObject"].ToString();

                var oppArtifact = JsonConvert.DeserializeObject<Opportunity>(opportunityJson.ToString(), new JsonSerializerSettings
                {
                    MissingMemberHandling = MissingMemberHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore
                });

                // Check access
                var roles = new List<Role>();
                roles.Add(new Role { DisplayName = "RelationshipManager" });
                var checkAccess = await _opportunityFactory.CheckAccessAsync(oppArtifact, roles, requestId);
                if (!checkAccess) _logger.LogError($"RequestId: {requestId} - CheckAccess DeleteItemAsync");

                if (oppArtifact.Metadata.OpportunityState == OpportunityState.Creating)
                {
                    var result = await _graphSharePointAppService.DeleteFileOrFolderAsync(_appOptions.ProposalManagementRootSiteId, $"TempFolder/{oppArtifact.DisplayName}", requestId);
                    // TODO: Check response
                }

                var json = await _graphSharePointAppService.DeleteListItemAsync(opportunitySiteList, id, requestId);
                Guard.Against.Null(json, nameof(json), requestId);

                return StatusCodes.Status204NoContent;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - OpportunityRepository_DeleteItemAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - OpportunityRepository_DeleteItemAsync Service Exception: {ex}");
            }
        }

        // Private methods
        private async Task<Opportunity> UpdateUsersAsync(Opportunity opportunity, string requestId = "")
        {
            _logger.LogInformation($"RequestId: {requestId} - OpportunityRepository_UpdateUsersAsync called.");

            try
            {
                Guard.Against.Null(opportunity, "OpportunityRepository_UpdateUsersAsync opportunity is null", requestId);

                var usersList = (await _userProfileRepository.GetAllAsync(requestId)).ToList();
                var teamMembers = opportunity.Content.TeamMembers.ToList();
                var updatedTeamMembers = new List<TeamMember>();
                
                foreach (var item in teamMembers)
                {
                    var updatedItem = TeamMember.Empty;
                    updatedItem.Id = item.Id;
                    updatedItem.DisplayName = item.DisplayName;
                    updatedItem.Status = item.Status;
                    updatedItem.AssignedRole = item.AssignedRole;
                    updatedItem.Fields = item.Fields;

                    var currMember = usersList.Find(x => x.Id == item.Id);

                    if (currMember != null)
                    {
                        updatedItem.DisplayName = currMember.DisplayName;
                        updatedItem.Fields = TeamMemberFields.Empty;
                        updatedItem.Fields.Mail = currMember.Fields.Mail;
                        updatedItem.Fields.Title = currMember.Fields.Title;
                        updatedItem.Fields.UserPrincipalName = currMember.Fields.UserPrincipalName;

                        var hasAssignedRole = currMember.Fields.UserRoles.Find(x => x.DisplayName == item.AssignedRole.DisplayName);

                        if (opportunity.Metadata.OpportunityState == OpportunityState.InProgress && hasAssignedRole != null)
                        {
                            updatedTeamMembers.Add(updatedItem);
                        }
                    }
                    else
                    {
                        if (opportunity.Metadata.OpportunityState != OpportunityState.InProgress)
                        {
                            updatedTeamMembers.Add(updatedItem);
                        }
                    }
                }
                opportunity.Content.TeamMembers = updatedTeamMembers;

                // TODO: Also update other users in opportunity like notes which has owner nd maps to a user profile

                return opportunity;
            }
            catch (Exception ex)
            {
                _logger.LogError($"RequestId: {requestId} - OpportunityRepository_UpdateUsersAsync Service Exception: {ex}");
                throw new ResponseException($"RequestId: {requestId} - OpportunityRepository_UpdateUsersAsync Service Exception: {ex}");
            }
        }
    }
}
