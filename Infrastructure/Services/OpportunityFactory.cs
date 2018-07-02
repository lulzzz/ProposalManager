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
	public class OpportunityFactory : BaseArtifactFactory<Opportunity>, IOpportunityFactory
	{
		private readonly GraphSharePointAppService _graphSharePointAppService;
		private readonly GraphUserAppService _graphUserAppService;
		private readonly INotificationRepository _notificationRepository;
		private readonly IUserProfileRepository _userProfileRepository;
        private readonly IRoleMappingRepository _roleMappingRepository;
        private readonly CardNotificationService _cardNotificationService;
        private readonly IUserContext _userContext;


		public OpportunityFactory(
			ILogger<OpportunityFactory> logger,
			IOptions<AppOptions> appOptions,
			GraphSharePointAppService graphSharePointAppService,
			GraphUserAppService graphUserAppService,
			INotificationRepository notificationRepository,
			IUserProfileRepository userProfileRepository,
            IRoleMappingRepository roleMappingRepository,
            CardNotificationService cardNotificationService,
            IUserContext userContext) : base(logger, appOptions)
		{
			Guard.Against.Null(graphSharePointAppService, nameof(graphSharePointAppService));
			Guard.Against.Null(graphUserAppService, nameof(graphUserAppService));
			Guard.Against.Null(notificationRepository, nameof(notificationRepository));
			Guard.Against.Null(userProfileRepository, nameof(userProfileRepository));
            Guard.Against.Null(roleMappingRepository, nameof(roleMappingRepository));
            Guard.Against.Null(cardNotificationService, nameof(cardNotificationService));
            Guard.Against.Null(userContext, nameof(userContext));
			_graphSharePointAppService = graphSharePointAppService;
			_graphUserAppService = graphUserAppService;
			_notificationRepository = notificationRepository;
			_userProfileRepository = userProfileRepository;
            _roleMappingRepository = roleMappingRepository;
            _cardNotificationService = cardNotificationService;
            _userContext = userContext;
		}

		public async Task<bool> CheckAccessAsync(Opportunity oppArtifact, List<Role> roles, string requestId = "")
		{
			try
			{
				Guard.Against.Null(oppArtifact, "OpportunityFactory_CheckAccess oppArtifact = null", requestId);
				Guard.Against.Null(roles, "OpportunityFactory_CheckAccess roles = null", requestId);

				var currentUser = (_userContext.User.Claims).ToList().Find(x => x.Type == "preferred_username")?.Value;
				Guard.Against.NullOrEmpty(currentUser, "OpportunityFactory_CheckAccess CurrentUser null-empty", requestId);

				var teamMember = (oppArtifact.Content.TeamMembers).ToList().Find(x => x.Fields.UserPrincipalName == currentUser);
				Guard.Against.Null(teamMember, "GetItemByIdAsync_CheckAccess_teamMember null", requestId);

                var userProfile = await _userProfileRepository.GetItemByUpnAsync(currentUser, requestId);
                Guard.Against.Null(userProfile, "GetItemByIdAsync_CheckAccess_userProfile null", requestId);

                foreach (var itm in roles)
				{
                    if (userProfile.Fields.UserRoles.Find(x => x.DisplayName == itm.DisplayName) == null)
					{
						// This role does not has access to opportunities
						_logger.LogError($"RequestId: {requestId} - OpportunityFactory_CheckAccess current user: {currentUser} AccessDeniedException");
						throw new AccessDeniedException($"RequestId: {requestId} - OpportunityFactory_CheckAccess current user: {currentUser} AccessDeniedException");
					}
				}

				return true;
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - OpportunityFactory_CheckAccess Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - OpportunityFactory_CheckAccess Service Exception: {ex}");
			}
		}

        public async Task<bool> CheckAccessAnyAsync(Opportunity oppArtifact, string requestId = "")
		{
			try
			{
                if (requestId.StartsWith("bot"))
                {
                    // TODO: Temp check for bot calls while bot sends token (currently is not)
                    return true;
                }

                Guard.Against.Null(oppArtifact, "OpportunityFactory_CheckAccessAny oppArtifact = null", requestId);

                var currentUser = (_userContext.User.Claims).ToList().Find(x => x.Type == "preferred_username")?.Value;
                Guard.Against.NullOrEmpty(currentUser, "OpportunityFactory_CheckAccessAny CurrentUser null-empty", requestId);

                var currUser = await _userProfileRepository.GetItemByUpnAsync(currentUser, requestId);
                if (currUser.Fields.UserRoles.Find(x => x.DisplayName == "Administrator") != null)
                {
                    return true;
                }

                var teamMember = (oppArtifact.Content.TeamMembers).ToList().Find(x => x.Fields.UserPrincipalName == currentUser);
                Guard.Against.Null(teamMember, "GetItemByIdAsync_CheckAccessAny_teamMember null", requestId);

                return true;
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - OpportunityFactory_CheckAccessAny Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - OpportunityFactory_CheckAccessAny Service Exception: {ex}");
			}
		}

        public async Task<Opportunity> CreateWorkflowAsync(Opportunity opportunity, string requestId = "")
		{
			try
			{
				// Set initial opportunity state
				opportunity.Metadata.OpportunityState = OpportunityState.Creating;

				// Remove empty sections from proposal document
				var porposalSectionList = new List<DocumentSection>();
				foreach (var item in opportunity.Content.ProposalDocument.Content.ProposalSectionList)
				{
					if (!String.IsNullOrEmpty(item.DisplayName))
					{
						porposalSectionList.Add(item);
					}
				}
				opportunity.Content.ProposalDocument.Content.ProposalSectionList = porposalSectionList;


				// Delete empty ChecklistItems
				opportunity.Content.Checklists = await RemoveEmptyFromChecklistAsync(opportunity.Content.Checklists, requestId);


				// Get Group id
				var opportunityName = opportunity.DisplayName.Replace("'", "");

				// Update status for team members & add them to team
				var isLoanOfficerSelected = false; //if true, relationship manager status should be completed
				var updatedTeamlist = new List<TeamMember>();
				foreach (var item in opportunity.Content.TeamMembers)
				{
					//var groupID = group;
					var userId = item.Id;
					var oItem = item;

					if (item.AssignedRole.DisplayName == "RelationshipManager")
					{
						oItem.Status = ActionStatus.InProgress;
					}
					else if (item.AssignedRole.DisplayName == "LoanOfficer")
					{
						if (!String.IsNullOrEmpty(item.Id))
						{
							isLoanOfficerSelected = true;
						}

                        // Enable the code below if team will be created before the opportunity
						//try
						//{
						//	Guard.Against.NullOrEmpty(userId, "CreateWorkflowAsync_LoanOffier_Ups Null or empty", requestId);
						//	var responseJson = await _graphUserAppService.AddGroupOwnerAsync(userId, groupID);
						//}
						//catch (Exception ex)
						//{
						//	_logger.LogError($"RequestId: {requestId} - userId: {userId} - AddGroupOwnerAsync error in CreateWorkflowAsync: {ex}");
						//}

					}
					else
					{
						// Nothing for other members
					}

					updatedTeamlist.Add(oItem);
				}

				opportunity.Content.TeamMembers = updatedTeamlist;

				// Update relationship manager status if loan officer has been selected
				if (isLoanOfficerSelected)
				{
					updatedTeamlist = new List<TeamMember>();
					foreach (var item in opportunity.Content.TeamMembers)
					{
						var prevItem = item;
						if (item.AssignedRole.DisplayName == "RelationshipManager")
						{
							prevItem.Status = ActionStatus.Completed;
						}
						updatedTeamlist.Add(prevItem);
					}

					opportunity.Content.TeamMembers = updatedTeamlist;
				}

				// Update note created by (if one) and set it to relationship manager
				if (opportunity.Content.Notes != null)
				{
					if (opportunity.Content.Notes?.Count > 0)
					{
						var currentUser = (_userContext.User.Claims).ToList().Find(x => x.Type == "preferred_username")?.Value;
						var callerUser = await _userProfileRepository.GetItemByUpnAsync(currentUser, requestId);

						if (callerUser != null)
						{
							opportunity.Content.Notes[0].CreatedBy = callerUser;
							opportunity.Content.Notes[0].CreatedDateTime = DateTimeOffset.Now;

						}
						else
						{
							_logger.LogWarning($"RequestId: {requestId} - CreateWorkflowAsync can't find {currentUser} to set note created by");
						}
					}
				}

                // Send notification
                // Define Sent To user profile
                var loanOfficer = opportunity.Content.TeamMembers.ToList().Find(x => x.AssignedRole.DisplayName == "LoanOfficer");             

                if (loanOfficer != null)
                {
                    try
                    {
                        _logger.LogInformation($"RequestId: {requestId} - CreateWorkflowAsync sendNotificationCardAsync new opportunity notification.");
                        var sendAccount = UserProfile.Empty;
                        sendAccount.Id = loanOfficer.Id;
                        sendAccount.DisplayName = loanOfficer.DisplayName;
                        sendAccount.Fields.UserPrincipalName = loanOfficer.Fields.UserPrincipalName;
                        var sendNotificationCard = await _cardNotificationService.sendNotificationCardAsync(opportunity, sendAccount, $"New opportunity {opportunity.DisplayName} has been assigned to ", requestId);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"RequestId: {requestId} - CreateWorkflowAsync sendNotificationCardAsync Action error: {ex}");
                    }
                }

                
                // wave 1 notifination
				if (loanOfficer != null)
				{
					var sendTo = loanOfficer.Fields.Mail ?? loanOfficer.Fields.UserPrincipalName;

					var relManager = opportunity.Content.TeamMembers.ToList().Find(x => x.AssignedRole.DisplayName == "RelationshipManager");

					if (relManager != null)
					{
						var notification = new Notification
						{
							Id = String.Empty,
							Title = "New opportunity '" + opportunity.DisplayName + "' has been assigned to you.",
							Fields = new NotificationFields
							{
								Message = "Please go to your dashboard or teams to view the new opportunity.",
								SentFrom = relManager.Fields.Mail ?? relManager.Fields.UserPrincipalName,
								SentTo = sendTo
							}
						};

						// ensure opportunity creation does not fail due to an error in creating and sending te notification
						try
						{
							var respCreateNotification = await _notificationRepository.CreateItemAsync(notification, requestId);
						}
						catch (Exception ex)
						{
							_logger.LogError($"RequestId: {requestId} - CreateWorkflowAsync CreateNotification Action error: {ex}");
						}
					}
				}
				

				return opportunity;
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - CreateWorkflowAsync Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - CreateWorkflowAsync Service Exception: {ex}");
			}
		}

        public async Task<Opportunity> UpdateWorkflowAsync(Opportunity opportunity, string requestId = "")
		{
			try
			{
                var initialState = opportunity.Metadata.OpportunityState;

				if (opportunity.Metadata.OpportunityState != OpportunityState.Creating)
				{
					if (opportunity.Content.CustomerDecision.Approved)
					{
						opportunity.Metadata.OpportunityState = OpportunityState.Accepted;
					}

					try
					{
						opportunity = await MoveTempFileToTeamAsync(opportunity, requestId);
					}
					catch(Exception ex)
					{
						_logger.LogError($"RequestId: {requestId} - UpdateWorkflowAsync_MoveTempFileToTeam Service Exception: {ex}");
					}

					// Add / update members
					// Get Group id
					//var opportunityName = opportunity.DisplayName.Replace(" ", "");
					var opportunityName = WebUtility.UrlEncode(opportunity.DisplayName);
					var options = new List<QueryParam>();

					options.Add(new QueryParam("filter", $"startswith(displayName,'{opportunityName}')"));

					var groupIdJson = await _graphUserAppService.GetGroupAsync(options, "", requestId);
					dynamic jsonDyn = groupIdJson;

					var group = String.Empty;
					if (groupIdJson.HasValues)
					{
						group = jsonDyn.value[0].id.ToString();
					}

					// add to team group
					var teamMembersComplete = 0;
					var isLoanOfficerSelected = false;
					foreach (var item in opportunity.Content.TeamMembers)
					{
						var groupID = group;
						var userId = item.Id;
						var oItem = item;

						if (item.AssignedRole.DisplayName == "RelationshipManager")
						{
							// In case an admin or background workflow will trigger this update after team/channels are created, relationship manager should also be added as owner
							try
							{
								Guard.Against.NullOrEmpty(item.Id, $"UpdateWorkflowAsync_{item.AssignedRole.DisplayName} Id NullOrEmpty", requestId);
								var responseJson = await _graphUserAppService.AddGroupOwnerAsync(userId, groupID, requestId);
							}
							catch (Exception ex)
							{
								_logger.LogError($"RequestId: {requestId} - userId: {userId} - UpdateWorkflowAsync_AddGroupOwnerAsync_{item.AssignedRole.DisplayName} error in CreateWorkflowAsync: {ex}");
							}
						}
						else if (item.AssignedRole.DisplayName == "LoanOfficer")
						{
							if (!String.IsNullOrEmpty(item.Id))
							{
								isLoanOfficerSelected = true; //Reltionship manager should be set to complete if loan officer is selected
							}
							try
							{
								Guard.Against.NullOrEmpty(item.Id, $"UpdateWorkflowAsync_{item.AssignedRole.DisplayName} Id NullOrEmpty", requestId);
								var responseJson = await _graphUserAppService.AddGroupOwnerAsync(userId, groupID, requestId);
							}
							catch (Exception ex)
							{
								_logger.LogError($"RequestId: {requestId} - userId: {userId} - UpdateWorkflowAsync_AddGroupOwnerAsync_{item.AssignedRole.DisplayName} error in CreateWorkflowAsync: {ex}");
							}

						}
						else
						{
							if (!String.IsNullOrEmpty(item.Fields.UserPrincipalName))
							{
								teamMembersComplete = teamMembersComplete + 1;
								try
								{
									Guard.Against.NullOrEmpty(item.Id, $"UpdateWorkflowAsync_{item.AssignedRole.DisplayName} Id NullOrEmpty", requestId);
									var responseJson = await _graphUserAppService.AddGroupMemberAsync(userId, groupID, requestId);
								}
								catch (Exception ex)
								{
									_logger.LogError($"RequestId: {requestId} - userId: {userId} - UpdateWorkflowAsync_AddGroupMemberAsync_{item.AssignedRole.DisplayName} error in CreateWorkflowAsync: {ex}");
								}
							}
						}
					}

					//Update status of team members
                    var oppCheckLists = opportunity.Content.Checklists.ToList();
                    var roleMappings = (await _roleMappingRepository.GetAllAsync(requestId)).ToList(); 

                    //TODO: LinQ
                    var updatedTeamlist = new List<TeamMember>();
					foreach (var item in opportunity.Content.TeamMembers)
					{
						var oItem = item;
						oItem.Status = ActionStatus.NotStarted;

						if (opportunity.Content.CustomerDecision.Approved)
						{
							oItem.Status = ActionStatus.Completed;
						}
						else
						{
                            var roleMap = roleMappings.Find(x => x.RoleName == item.AssignedRole.DisplayName);

                            if (item.AssignedRole.DisplayName != "LoanOfficer" && item.AssignedRole.DisplayName != "RelationshipManager")
                            {
                                if (roleMap != null)
                                {
                                    _logger.LogInformation($"RequestId: {requestId} - UpdateOpportunityAsync teamMember status sync with checklist status RoleName: {roleMap.RoleName}");

                                    var checklistItm = oppCheckLists.Find(x => x.ChecklistChannel == roleMap.Channel);
                                    if (checklistItm != null)
                                    {
                                        _logger.LogInformation($"RequestId: {requestId} - UpdateOpportunityAsync teamMember status sync with checklist status: {checklistItm.ChecklistStatus.Name}");
                                        oItem.Status = checklistItm.ChecklistStatus;
                                    }
                                }
                            }
                            else if (item.AssignedRole.DisplayName == "RelationshipManager")
                            {
                                var exisitngLoanOfficers = ((opportunity.Content.TeamMembers).ToList()).Find(x => x.AssignedRole.DisplayName == "LoanOfficer");
                                if (exisitngLoanOfficers == null)
                                {
                                    oItem.Status = ActionStatus.InProgress;
                                }
                                else
                                {
                                    oItem.Status = ActionStatus.Completed;
                                }
                            }
                            else if (item.AssignedRole.DisplayName == "LoanOfficer")
                            {
                                var teamList = ((opportunity.Content.TeamMembers).ToList()).FindAll(x => x.AssignedRole.DisplayName != "LoanOfficer" && x.AssignedRole.DisplayName != "RelationshipManager");
                                if (teamList != null)
                                {
                                    var expectedTeam = roleMappings.FindAll(x => x.RoleName != "LoanOfficer" && x.RoleName != "RelationshipManager" && x.RoleName != "Administrator");
                                    if (expectedTeam != null)
                                    {
                                        if (teamList.Count != 0)
                                        {
                                            oItem.Status = ActionStatus.InProgress;
                                        }
                                        if (teamList.Count >= expectedTeam.Count)
                                        {
                                            oItem.Status = ActionStatus.Completed;
                                        }
                                    }
                                }
                            }
                        }

                        updatedTeamlist.Add(oItem);
					}
                    
                    opportunity.Content.TeamMembers = updatedTeamlist;
				}

                // Send notification
                _logger.LogInformation($"RequestId: {requestId} - UpdateWorkflowAsync initialState: {initialState.Name} - {opportunity.Metadata.OpportunityState.Name}");
                if (initialState.Value != opportunity.Metadata.OpportunityState.Value)
                {
                    try
                    {
                        _logger.LogInformation($"RequestId: {requestId} - CreateWorkflowAsync sendNotificationCardAsync opportunity state change notification.");
                        var sendTo = UserProfile.Empty;
                        var sendNotificationCard = await _cardNotificationService.sendNotificationCardAsync(opportunity, sendTo, $"Opportunity state for {opportunity.DisplayName} has been changed to {opportunity.Metadata.OpportunityState.Name}", requestId);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"RequestId: {requestId} - CreateWorkflowAsync sendNotificationCardAsync OpportunityState error: {ex}");
                    }
                }

                // Delete empty ChecklistItems
                //opportunity.Content.Checklists = await RemoveEmptyFromChecklist(opportunity.Content.Checklists, requestId);

                return opportunity;
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - UpdateWorkflowAsync Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - UpdateWorkflowAsync Service Exception: {ex}");
			}
		}

        // Workflow Actions
        public async Task<Opportunity> MoveTempFileToTeamAsync(Opportunity opportunity, string requestId = "")
		{
			try
			{

				// Find entries that need to be moved
				var moveFiles = false;
				foreach(var itm in opportunity.DocumentAttachments)
				{
					if (itm.DocumentUri == "TempFolder") moveFiles = true;
				}

				if (moveFiles)
				{
					var fromSiteId = _appOptions.ProposalManagementRootSiteId;
					var toSiteId = String.Empty;
					var fromItemPath = String.Empty;
					var toItemPath = String.Empty;

					string pattern = @"[ `~!@#$%^&*()_|+\-=?;:'" + '"' + @",.<>\{\}\[\]\\\/]";
					string replacement = "";

					Regex regEx = new Regex(pattern);
					var path = regEx.Replace(opportunity.DisplayName, replacement);
					//var path = WebUtility.UrlEncode(opportunity.DisplayName);
					//var path = opportunity.DisplayName.Replace(" ", "");

					var siteIdResponse = await _graphSharePointAppService.GetSiteIdAsync(_appOptions.SharePointHostName, path, requestId);
					dynamic responseDyn = siteIdResponse;
					toSiteId = responseDyn.id.ToString();

					if (!String.IsNullOrEmpty(toSiteId))
					{
						var updatedDocumentAttachments = new List<DocumentAttachment>();
						foreach (var itm in opportunity.DocumentAttachments)
						{
							var updDoc = DocumentAttachment.Empty;
							if (itm.DocumentUri == "TempFolder")
							{
								fromItemPath = $"TempFolder/{opportunity.DisplayName}/{itm.FileName}";
								toItemPath = $"General/{itm.FileName}";

								var resp = new JObject();
								try
								{
									resp = await _graphSharePointAppService.MoveFileAsync(fromSiteId, fromItemPath, toSiteId, toItemPath, requestId);
									updDoc.Id = new Guid().ToString();
									updDoc.DocumentUri = String.Empty;
									//doc.Id = resp.id;
								}
								catch (Exception ex)
								{
									_logger.LogWarning($"RequestId: {requestId} - MoveTempFileToTeam: from: {fromItemPath} to: {toItemPath} Service Exception: {ex}");
								}
							}

							updDoc.FileName = itm.FileName;
							updDoc.Note = itm.Note ?? String.Empty;
							updDoc.Tags = itm.Tags ?? String.Empty;
							updDoc.Category = Category.Empty;
							updDoc.Category.Id = itm.Category.Id;
							updDoc.Category.Name = itm.Category.Name;

							updatedDocumentAttachments.Add(updDoc);
						}

						opportunity.DocumentAttachments = updatedDocumentAttachments;

						// Delete temp files
						var result = await _graphSharePointAppService.DeleteFileOrFolderAsync(_appOptions.ProposalManagementRootSiteId, $"TempFolder/{opportunity.DisplayName}", requestId);

					}
				}

				return opportunity;
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - MoveTempFileToTeam Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - MoveTempFileToTeam Service Exception: {ex}");
			}
		}

        public Task<IList<Checklist>> RemoveEmptyFromChecklistAsync(IList<Checklist> checklists, string requestId = "")
		{
			try
			{
				var newChecklists = new List<Checklist>();
				foreach (var item in checklists)
				{
					var newChecklist = new Checklist();
					newChecklist.ChecklistTaskList = new List<ChecklistTask>();
					newChecklist.ChecklistChannel = item.ChecklistChannel;
					newChecklist.ChecklistStatus = item.ChecklistStatus;
					newChecklist.Id = item.Id;

					
					foreach (var sItem in item.ChecklistTaskList)
					{
						var newChecklistTask = new ChecklistTask();
						if (!String.IsNullOrEmpty(sItem.Id) && !String.IsNullOrEmpty(sItem.ChecklistItem))
						{
							newChecklistTask.Id = sItem.Id;
							newChecklistTask.ChecklistItem = sItem.ChecklistItem;
							newChecklistTask.Completed = sItem.Completed;
							newChecklistTask.FileUri = sItem.FileUri;

							newChecklist.ChecklistTaskList.Add(newChecklistTask);
						}
					}

					newChecklists.Add(newChecklist);
				}

				return Task.FromResult<IList<Checklist>>(newChecklists);
			}
			catch (Exception ex)
			{
				_logger.LogError($"RequestId: {requestId} - RemoveEmptyFromChecklist Service Exception: {ex}");
				throw new ResponseException($"RequestId: {requestId} - RemoveEmptyFromChecklist Service Exception: {ex}");
			}
		}
	}
}
