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
    public class OpportunityHelpers
    {
        protected readonly ILogger _logger;
        protected readonly AppOptions _appOptions;
        private readonly UserProfileHelpers _userProfileHelpers;
        private readonly IRoleMappingRepository _roleMappingRepository;
        private readonly CardNotificationService _cardNotificationService;

        /// <summary>
        /// Constructor
        /// </summary>
        public OpportunityHelpers(
            ILogger<UserProfileHelpers> logger,
            IOptions<AppOptions> appOptions,
            UserProfileHelpers userProfileHelpers,
            IRoleMappingRepository roleMappingRepository,
            CardNotificationService cardNotificationService)
        {
            Guard.Against.Null(logger, nameof(logger));
            Guard.Against.Null(appOptions, nameof(appOptions));
            Guard.Against.Null(userProfileHelpers, nameof(userProfileHelpers));
            Guard.Against.Null(roleMappingRepository, nameof(roleMappingRepository));
            Guard.Against.Null(cardNotificationService, nameof(cardNotificationService));

            _logger = logger;
            _appOptions = appOptions.Value;
            _userProfileHelpers = userProfileHelpers;
            _roleMappingRepository = roleMappingRepository;
            _cardNotificationService = cardNotificationService;
        }

        public async Task<OpportunityViewModel> ToOpportunityViewModelAsync(Opportunity opportunity, string requestId = "")
        {
            return await OpportunityToViewModelAsync(opportunity, requestId);
        }

        public async Task<OpportunityViewModel> OpportunityToViewModelAsync(Opportunity entity, string requestId = "")
        {
            var oppId = entity.Id;
            try
            {
                //var entityDto = TinyMapper.Map<OpportunityViewModel>(entity);
                var viewModel = new OpportunityViewModel
                {
                    Id = entity.Id,
                    DisplayName = entity.DisplayName,
                    Reference = entity.Reference,
                    Version = entity.Version,
                    OpportunityState = OpportunityStateModel.FromValue(entity.Metadata.OpportunityState.Value),
                    DealSize = entity.Metadata.DealSize,
                    AnnualRevenue = entity.Metadata.AnnualRevenue,
                    OpenedDate = entity.Metadata.OpenedDate,
                    Industry = new IndustryModel
                    {
                        Name = entity.Metadata.Industry.Name,
                        Id = entity.Metadata.Industry.Id
                    },
                    Region = new RegionModel
                    {
                        Name = entity.Metadata.Region.Name,
                        Id = entity.Metadata.Region.Id
                    },
                    Margin = entity.Metadata.Margin,
                    Rate = entity.Metadata.Rate,
                    DebtRatio = entity.Metadata.DebtRatio,
                    Purpose = entity.Metadata.Purpose,
                    DisbursementSchedule = entity.Metadata.DisbursementSchedule,
                    CollateralAmount = entity.Metadata.CollateralAmount,
                    Guarantees = entity.Metadata.Guarantees,
                    RiskRating = entity.Metadata.RiskRating,
                    OpportunityChannelId = entity.Metadata.OpportunityChannelId,
                    Customer = new CustomerModel
                    {
                        DisplayName = entity.Metadata.Customer.DisplayName,
                        Id = entity.Metadata.Customer.Id,
                        ReferenceId = entity.Metadata.Customer.ReferenceId
                    },
                    TeamMembers = new List<TeamMemberModel>(),
                    Notes = new List<NoteModel>(),
                    Checklists = new List<ChecklistModel>(),
                    CustomerDecision = new CustomerDecisionModel
                    {
                        Id = entity.Content.CustomerDecision.Id,
                        Approved = entity.Content.CustomerDecision.Approved,
                        ApprovedDate = entity.Content.CustomerDecision.ApprovedDate,
                        LoanDisbursed = entity.Content.CustomerDecision.LoanDisbursed
                    }
                };

                viewModel.ProposalDocument = new ProposalDocumentModel();
                viewModel.ProposalDocument.Id = entity.Content.ProposalDocument.Id;
                viewModel.ProposalDocument.DisplayName = entity.Content.ProposalDocument.DisplayName;
                viewModel.ProposalDocument.Reference = entity.Content.ProposalDocument.Reference;
                viewModel.ProposalDocument.DocumentUri = entity.Content.ProposalDocument.Metadata.DocumentUri;
                viewModel.ProposalDocument.Category = new CategoryModel();
                viewModel.ProposalDocument.Category.Id = entity.Content.ProposalDocument.Metadata.Category.Id;
                viewModel.ProposalDocument.Category.Name = entity.Content.ProposalDocument.Metadata.Category.Name;
                viewModel.ProposalDocument.Content = new ProposalDocumentContentModel();
                viewModel.ProposalDocument.Content.ProposalSectionList = new List<DocumentSectionModel>();
                viewModel.ProposalDocument.Notes = new List<NoteModel>();

                viewModel.ProposalDocument.Tags = entity.Content.ProposalDocument.Metadata.Tags;
                viewModel.ProposalDocument.Version = entity.Content.ProposalDocument.Version;


                // Checklists
                foreach (var item in entity.Content.Checklists)
                {
                    var checklistTasks = new List<ChecklistTaskModel>();
                    foreach (var subitem in item.ChecklistTaskList)
                    {
                        var checklistItem = new ChecklistTaskModel
                        {
                            Id = subitem.Id,
                            ChecklistItem = subitem.ChecklistItem,
                            Completed = subitem.Completed,
                            FileUri = subitem.FileUri
                        };
                        checklistTasks.Add(checklistItem);
                    }

                    var checklistModel = new ChecklistModel
                    {
                        Id = item.Id,
                        ChecklistStatus = item.ChecklistStatus,
                        ChecklistTaskList = checklistTasks,
                        ChecklistChannel = item.ChecklistChannel
                    };
                    viewModel.Checklists.Add(checklistModel);
                }

                
                // TeamMembers
                foreach (var item in entity.Content.TeamMembers.ToList())
                {
                    var memberModel = new TeamMemberModel();
                    memberModel.Status = item.Status;
                    memberModel.AssignedRole = await _userProfileHelpers.RoleToViewModelAsync(item.AssignedRole, requestId);
                    memberModel.Id = item.Id;
                    memberModel.DisplayName = item.DisplayName;
                    memberModel.Mail = item.Fields.Mail;
                    memberModel.UserPrincipalName = item.Fields.UserPrincipalName;
                    memberModel.Title = item.Fields.Title ?? String.Empty;

                    viewModel.TeamMembers.Add(memberModel);
                }


                // Notes
                foreach (var item in entity.Content.Notes.ToList())
                {
                    var note = new NoteModel();
                    note.Id = item.Id;

                    var userProfile = new UserProfileViewModel();
                    userProfile.Id = item.CreatedBy.Id;
                    userProfile.DisplayName = item.CreatedBy.DisplayName;
                    userProfile.Mail = item.CreatedBy.Fields.Mail;
                    userProfile.UserPrincipalName = item.CreatedBy.Fields.UserPrincipalName;
                    userProfile.UserRoles = await _userProfileHelpers.RolesToViewModelAsync(item.CreatedBy.Fields.UserRoles, requestId);

                    note.CreatedBy = userProfile;
                    note.NoteBody = item.NoteBody;
                    note.CreatedDateTime = item.CreatedDateTime;

                    viewModel.Notes.Add(note);
                }


                // ProposalDocument Notes
                foreach (var item in entity.Content.ProposalDocument.Metadata.Notes.ToList())
                {
                    var docNote = new NoteModel();

                    docNote.Id = item.Id;
                    docNote.CreatedDateTime = item.CreatedDateTime;
                    docNote.NoteBody = item.NoteBody;
                    docNote.CreatedBy = new UserProfileViewModel
                    {
                        Id = item.CreatedBy.Id,
                        DisplayName = item.CreatedBy.DisplayName,
                        Mail = item.CreatedBy.Fields.Mail,
                        UserPrincipalName = item.CreatedBy.Fields.UserPrincipalName,
                        UserRoles = await _userProfileHelpers.RolesToViewModelAsync(item.CreatedBy.Fields.UserRoles, requestId)
                    };

                    viewModel.ProposalDocument.Notes.Add(docNote);
                }


                // ProposalDocument ProposalSectionList
                foreach (var item in entity.Content.ProposalDocument.Content.ProposalSectionList.ToList())
                {
                    if (!String.IsNullOrEmpty(item.Id))
                    {
                        var docSectionModel = new DocumentSectionModel();
                        docSectionModel.Id = item.Id;
                        docSectionModel.DisplayName = item.DisplayName;
                        docSectionModel.LastModifiedDateTime = item.LastModifiedDateTime;
                        docSectionModel.Owner = new UserProfileViewModel();
                        if (item.Owner != null)
                        {
                            docSectionModel.Owner.Id = item.Owner.Id ?? String.Empty;
                            docSectionModel.Owner.DisplayName = item.Owner.DisplayName ?? String.Empty;
                            if (item.Owner.Fields != null)
                            {
                                docSectionModel.Owner.Mail = item.Owner.Fields.Mail ?? String.Empty;
                                docSectionModel.Owner.UserPrincipalName = item.Owner.Fields.UserPrincipalName ?? String.Empty;
                                docSectionModel.Owner.UserRoles = new List<RoleModel>();

                                if (item.Owner.Fields.UserRoles != null)
                                {
                                    docSectionModel.Owner.UserRoles = await _userProfileHelpers.RolesToViewModelAsync(item.Owner.Fields.UserRoles, requestId);
                                }
                            }
                            else
                            {
                                docSectionModel.Owner.Mail = String.Empty;
                                docSectionModel.Owner.UserPrincipalName = String.Empty;
                                docSectionModel.Owner.UserRoles = new List<RoleModel>();

                                if (item.Owner.Fields.UserRoles != null)
                                {
                                    docSectionModel.Owner.UserRoles = await _userProfileHelpers.RolesToViewModelAsync(item.Owner.Fields.UserRoles, requestId);
                                }
                            }
                        }

                        docSectionModel.SectionStatus = item.SectionStatus;
                        docSectionModel.SubSectionId = item.SubSectionId;
                        docSectionModel.AssignedTo = new UserProfileViewModel();
                        if (item.AssignedTo != null)
                        {
                            docSectionModel.AssignedTo.Id = item.AssignedTo.Id;
                            docSectionModel.AssignedTo.DisplayName = item.AssignedTo.DisplayName;
                            docSectionModel.AssignedTo.Mail = item.AssignedTo.Fields.Mail;
                            docSectionModel.AssignedTo.Title = item.AssignedTo.Fields.Title;
                            docSectionModel.AssignedTo.UserPrincipalName = item.AssignedTo.Fields.UserPrincipalName;
                            // TODO: Not including role info since it is not relevant but if needed it needs to be set here
                        }
                        docSectionModel.Task = item.Task;

                        viewModel.ProposalDocument.Content.ProposalSectionList.Add(docSectionModel);
                    }
                }

                // DocumentAttachments
                viewModel.DocumentAttachments = new List<DocumentAttachmentModel>();
                if (entity.DocumentAttachments != null)
                {
                    foreach (var itm in entity.DocumentAttachments)
                    {
                        var doc = new DocumentAttachmentModel();
                        doc.Id = itm.Id ?? String.Empty;
                        doc.FileName = itm.FileName ?? String.Empty;
                        doc.Note = itm.Note ?? String.Empty;
                        doc.Tags = itm.Tags ?? String.Empty;
                        doc.Category = new CategoryModel();
                        doc.Category.Id = itm.Category.Id;
                        doc.Category.Name = itm.Category.Name;
                        doc.DocumentUri = itm.DocumentUri;

                        viewModel.DocumentAttachments.Add(doc);
                    }
                }

                return viewModel;
            }
            catch (Exception ex)
            {
                // TODO: _logger.LogError("MapToViewModelAsync error: " + ex);
                throw new ResponseException($"RequestId: {requestId} - OpportunityToViewModelAsync oppId: {oppId} - failed to map opportunity: {ex}");
            }
        }

        public async Task<Opportunity> ToOpportunityAsync(OpportunityViewModel model, Opportunity opportunity, string requestId = "")
        {
            return await OpportunityToEntityAsync(model, opportunity, requestId);
        }

        #region MAP: model -> entity
        private async Task<Opportunity> OpportunityToEntityAsync(OpportunityViewModel viewModel, Opportunity opportunity, string requestId = "")
        {
            var oppId = viewModel.Id;

            try
            {
                var entity = opportunity;
                var entityIsEmpty = true;
                if (!String.IsNullOrEmpty(entity.DisplayName)) entityIsEmpty = false; // If empty we should not send any notifications since it is just a reference opportunity schema

                entity.Id = viewModel.Id ?? String.Empty;
                entity.DisplayName = viewModel.DisplayName ?? String.Empty;
                entity.Reference = viewModel.Reference ?? String.Empty;
                entity.Version = viewModel.Version ?? String.Empty;

                // DocumentAttachments
                if (entity.DocumentAttachments == null) entity.DocumentAttachments = new List<DocumentAttachment>();
                if (viewModel.DocumentAttachments != null)
                {
                    var newDocumentAttachments = new List<DocumentAttachment>();
                    foreach (var itm in viewModel.DocumentAttachments)
                    {
                        var doc = entity.DocumentAttachments.ToList().Find(x => x.Id == itm.Id);
                        if (doc == null)
                        {
                            doc = DocumentAttachment.Empty;
                        }

                        doc.Id = itm.Id;
                        doc.FileName = itm.FileName ?? String.Empty;
                        doc.DocumentUri = itm.DocumentUri ?? String.Empty;
                        doc.Category = Category.Empty;
                        doc.Category.Id = itm.Category.Id ?? String.Empty;
                        doc.Category.Name = itm.Category.Name ?? String.Empty;
                        doc.Tags = itm.Tags ?? String.Empty;
                        doc.Note = itm.Note ?? String.Empty;

                        newDocumentAttachments.Add(doc);
                    }

                    // TODO: P2 create logic for replace and support for other artifact types for now we replace the whole list
                    entity.DocumentAttachments = newDocumentAttachments;
                }

                // Content
                if (entity.Content == null) entity.Content = OpportunityContent.Empty;


                // Checklists
                if (entity.Content.Checklists == null) entity.Content.Checklists = new List<Checklist>();
                if (viewModel.Checklists != null)
                {
                    // List of checklists that status changed thus team members need to be sent with a notification
                    var statusChangedChecklists = new List<Checklist>();

                    var updatedList = new List<Checklist>();
                    // LIST: Content/CheckList/ChecklistTaskList
                    foreach (var item in viewModel.Checklists)
                    {
                        var checklist = Checklist.Empty;
                        var existinglist = entity.Content.Checklists.ToList().Find(x => x.Id == item.Id);
                        if (existinglist != null) checklist = existinglist;

                        var addToChangedList = false;
                        if (checklist.ChecklistStatus.Value != item.ChecklistStatus.Value)
                        {
                            addToChangedList = true;
                        }

                        checklist.Id = item.Id ?? String.Empty;
                        checklist.ChecklistStatus = ActionStatus.FromValue(item.ChecklistStatus.Value);
                        checklist.ChecklistTaskList = new List<ChecklistTask>();
                        checklist.ChecklistChannel = item.ChecklistChannel ?? String.Empty;

                        foreach (var subitem in item.ChecklistTaskList)
                        {
                            var checklistTask = new ChecklistTask
                            {
                                Id = subitem.Id ?? String.Empty,
                                ChecklistItem = subitem.ChecklistItem ?? String.Empty,
                                Completed = subitem.Completed,
                                FileUri = subitem.FileUri ?? String.Empty
                            };
                            checklist.ChecklistTaskList.Add(checklistTask);
                        }

                        // Add checklist for notifications, notification is sent below during teamMembers iterations
                        if (addToChangedList)
                        {
                            statusChangedChecklists.Add(checklist);
                        }

                        updatedList.Add(checklist);
                    }

                    // Send notifications for changed checklists
                    if (statusChangedChecklists.Count > 0 && !entityIsEmpty)
                    {
                        try
                        {
                            if (statusChangedChecklists.Count > 0)
                            {
                                var checkLists = String.Empty;
                                foreach (var chkItm in statusChangedChecklists)
                                {
                                    checkLists = checkLists + $"'{chkItm.ChecklistChannel}' ";
                                }

                                var sendToList = new List<UserProfile>();
                                if (!String.IsNullOrEmpty(viewModel.OpportunityChannelId)) entity.Metadata.OpportunityChannelId = viewModel.OpportunityChannelId;

                                _logger.LogInformation($"RequestId: {requestId} - UpdateWorkflowAsync sendNotificationCardAsync checklist status changed notification. Number of hecklists: {statusChangedChecklists.Count}");
                                var sendNotificationCard = await _cardNotificationService.sendNotificationCardAsync(entity, sendToList, $"Status updated for opportunity checklist(s): {checkLists} ", requestId);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError($"RequestId: {requestId} - UpdateWorkflowAsync sendNotificationCardAsync checklist status change error: {ex}");
                        }
                    }

                    entity.Content.Checklists = updatedList;
                }

                if (entity.Content.Checklists.Count == 0)
                {
                    // Checklist empty create a default set

                    var roleMappingList = (await _roleMappingRepository.GetAllAsync(requestId)).ToList();

                    foreach(var item in roleMappingList)
                    {
                        if (item.ProcessType.ToLower() == "checklisttab")
                        {
                            var checklist = new Checklist
                            {
                                Id = item.ProcessStep,
                                ChecklistChannel = item.Channel,
                                ChecklistStatus = ActionStatus.NotStarted,
                                ChecklistTaskList = new List<ChecklistTask>()
                            };
                            entity.Content.Checklists.Add(checklist);
                        }
                    }
                }


                // CustomerDecision
                if (entity.Content.CustomerDecision == null) entity.Content.CustomerDecision = CustomerDecision.Empty;
                if (viewModel.CustomerDecision != null)
                {
                    entity.Content.CustomerDecision.Id = viewModel.CustomerDecision.Id ?? String.Empty;
                    entity.Content.CustomerDecision.Approved = viewModel.CustomerDecision.Approved;
                    if (viewModel.CustomerDecision.ApprovedDate != null) entity.Content.CustomerDecision.ApprovedDate = viewModel.CustomerDecision.ApprovedDate;
                    if (viewModel.CustomerDecision.LoanDisbursed != null) entity.Content.CustomerDecision.LoanDisbursed = viewModel.CustomerDecision.LoanDisbursed;
                }


                // LIST: Content/Notes
                if (entity.Content.Notes == null) entity.Content.Notes = new List<Note>();
                if (viewModel.Notes != null)
                {
                    var updatedNotes = entity.Content.Notes.ToList();
                    foreach (var item in viewModel.Notes)
                    {
                        var note = updatedNotes.Find(itm => itm.Id == item.Id);
                        if (note != null)
                        {
                            updatedNotes.Remove(note);
                        }
                        updatedNotes.Add(await NoteToEntityAsync(item, requestId));
                    }

                    entity.Content.Notes = updatedNotes;
                }


                // TeamMembers
                if (entity.Content.TeamMembers == null) entity.Content.TeamMembers = new List<TeamMember>();
                if (viewModel.TeamMembers != null)
                {
                    var updatedTeamMembers = new List<TeamMember>();

                    // Update team members
                    foreach (var item in viewModel.TeamMembers)
                    {
                        updatedTeamMembers.Add(await TeamMemberToEntityAsync(item));
                    }
                    entity.Content.TeamMembers = updatedTeamMembers;
                }

                // ProposalDocument
                if (entity.Content.ProposalDocument == null) entity.Content.ProposalDocument = ProposalDocument.Empty;
                if (viewModel.ProposalDocument != null) entity.Content.ProposalDocument = await ProposalDocumentToEntityAsync(viewModel, entity.Content.ProposalDocument, requestId);

                // Metadata
                if (entity.Metadata == null) entity.Metadata = OpportunityMetadata.Empty;
                entity.Metadata.AnnualRevenue = viewModel.AnnualRevenue;
                entity.Metadata.CollateralAmount = viewModel.CollateralAmount;

                if (entity.Metadata.Customer == null) entity.Metadata.Customer = Customer.Empty;
                entity.Metadata.Customer.DisplayName = viewModel.Customer.DisplayName ?? String.Empty;
                entity.Metadata.Customer.Id = viewModel.Customer.Id ?? String.Empty;
                entity.Metadata.Customer.ReferenceId = viewModel.Customer.ReferenceId ?? String.Empty;

                entity.Metadata.DealSize = viewModel.DealSize;
                entity.Metadata.DebtRatio = viewModel.DebtRatio;
                entity.Metadata.DisbursementSchedule = viewModel.DisbursementSchedule ?? String.Empty;
                entity.Metadata.Guarantees = viewModel.Guarantees ?? String.Empty;

                if (entity.Metadata.Industry == null) entity.Metadata.Industry = new Industry();
                if (viewModel.Industry != null) entity.Metadata.Industry = await IndustryToEntityAsync(viewModel.Industry);

                entity.Metadata.Margin = viewModel.Margin;

                if (entity.Metadata.OpenedDate == null) entity.Metadata.OpenedDate = DateTimeOffset.MinValue;
                if (viewModel.OpenedDate != null) entity.Metadata.OpenedDate = viewModel.OpenedDate;

                if (entity.Metadata.OpportunityState == null) entity.Metadata.OpportunityState = OpportunityState.Creating;
                if (viewModel.OpportunityState != null) entity.Metadata.OpportunityState = OpportunityState.FromValue(viewModel.OpportunityState.Value);

                entity.Metadata.Purpose = viewModel.Purpose ?? String.Empty;
                entity.Metadata.Rate = viewModel.Rate;

                if (entity.Metadata.Region == null) entity.Metadata.Region = Region.Empty;
                if (viewModel.Region != null) entity.Metadata.Region = await RegionToEntityAsync(viewModel.Region);

                entity.Metadata.RiskRating = viewModel.RiskRating;

                // if to avoid deleting channelId if vieModel passes empty and a value was already in opportunity
                if (!String.IsNullOrEmpty(viewModel.OpportunityChannelId)) entity.Metadata.OpportunityChannelId = viewModel.OpportunityChannelId;

                return entity;
            }
            catch (Exception ex)
            {
                //_logger.LogError("MapFromViewModelAsync error: " + ex);
                throw new ResponseException($"RequestId: {requestId} - OpportunityToEntityAsync oppId: {oppId} - failed to map opportunity: {ex}");
            }
        }

        private async Task<ProposalDocument> ProposalDocumentToEntityAsync(OpportunityViewModel viewModel, ProposalDocument proposalDocument, string requestId = "")
        {
            try
            {
                Guard.Against.Null(proposalDocument, "ProposalDocumentToEntityAsync", requestId);
                var entity = proposalDocument;
                var model = viewModel.ProposalDocument;

                entity.Id = model.Id ?? String.Empty;
                entity.DisplayName = model.DisplayName ?? String.Empty;
                entity.Reference = model.Reference ?? String.Empty;
                entity.Version = model.Version ?? String.Empty;

                if (entity.Content == null)
                {
                    entity.Content = ProposalDocumentContent.Empty;
                }

                
                // Storing previous section lists to compare and trigger notification if assigment changes
                var currProposalSectionList = entity.Content.ProposalSectionList.ToList();

                // Proposal sections are always overwritten
                entity.Content.ProposalSectionList = new List<DocumentSection>();

                if (model.Content.ProposalSectionList != null)
                {
                    var OwnerSendList = new List<UserProfile>(); //receipients list for notifications
                    var AssignedToSendList = new List<UserProfile>(); //receipients list for notifications
                    // LIST: ProposalSectionList
                    foreach (var item in model.Content.ProposalSectionList)
                    {
                        var documentSection = new DocumentSection();
                        documentSection.DisplayName = item.DisplayName ?? String.Empty;
                        documentSection.Id = item.Id ?? String.Empty;
                        documentSection.LastModifiedDateTime = item.LastModifiedDateTime;
                        documentSection.Owner = await _userProfileHelpers.UserProfileToEntityAsync(item.Owner ?? new UserProfileViewModel(), requestId);
                        documentSection.SectionStatus = ActionStatus.FromValue(item.SectionStatus.Value);
                        documentSection.SubSectionId = item.SubSectionId ?? String.Empty;
                        documentSection.AssignedTo = await _userProfileHelpers.UserProfileToEntityAsync(item.AssignedTo ?? new UserProfileViewModel(), requestId);
                        documentSection.Task = item.Task ?? String.Empty;

                        // Check values to see if notification trigger is needed
                        var prevSectionList = currProposalSectionList.Find(x => x.Id == documentSection.Id);
                        if (prevSectionList != null)
                        {
                            if (prevSectionList.Owner.Id != documentSection.Owner.Id)
                            {
                                OwnerSendList.Add(documentSection.Owner);
                            }

                            if (prevSectionList.AssignedTo.Id != documentSection.AssignedTo.Id)
                            {
                                AssignedToSendList.Add(documentSection.AssignedTo);
                            }
                        }

                        entity.Content.ProposalSectionList.Add(documentSection);
                    }

                    // AssignedToSendList notifications
                    // Section owner changed / assigned
                    try
                    {
                        if (OwnerSendList.Count > 0)
                        {
                            _logger.LogInformation($"RequestId: {requestId} - ProposalDocumentToEntityAsync sendNotificationCardAsync for owner changed notification.");
                            var notificationOwner = await _cardNotificationService.sendNotificationCardAsync(
                                viewModel.DisplayName,
                                viewModel.OpportunityChannelId,
                                OwnerSendList,
                                $"Section(s) in the proposal document for opportunity {viewModel.DisplayName} has new/updated owners ",
                                requestId);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"RequestId: {requestId} - ProposalDocumentToEntityAsync sendNotificationCardAsync for owner error: {ex}");
                    }

                    // Section AssignedTo changed / assigned
                    try
                    {
                        if (AssignedToSendList.Count > 0)
                        {
                            _logger.LogInformation($"RequestId: {requestId} - ProposalDocumentToEntityAsync sendNotificationCardAsync for AssigedTo changed notification.");
                            var notificationAssignedTo = await _cardNotificationService.sendNotificationCardAsync(
                                viewModel.DisplayName,
                                viewModel.OpportunityChannelId,
                                AssignedToSendList,
                                $"Task(s) in the proposal document for opportunity {viewModel.DisplayName} has new/updated assigments ",
                                requestId);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"RequestId: {requestId} - ProposalDocumentToEntityAsync sendNotificationCardAsync for AssignedTo error: {ex}");
                    }
                }


                // Metadata
                if (entity.Metadata == null)
                {
                    entity.Metadata = DocumentMetadata.Empty;
                }

                entity.Metadata.DocumentUri = model.DocumentUri;
                entity.Metadata.Tags = model.Tags;
                if (entity.Metadata.Category == null)
                {
                    entity.Metadata.Category = new Category();
                }

                entity.Metadata.Category.Id = model.Category.Id ?? String.Empty;
                entity.Metadata.Category.Name = model.Category.Name ?? String.Empty;

                if (entity.Metadata.Notes == null)
                {
                    entity.Metadata.Notes = new List<Note>();
                }

                if (model.Notes != null)
                {
                    // LIST: Notes
                    foreach (var item in model.Notes)
                    {
                        entity.Metadata.Notes.Add(await NoteToEntityAsync(item, requestId));
                    }
                }

                return entity;
            }
            catch (Exception ex)
            {
                //_logger.LogError(ex.Message);
                throw ex;
            }
        }

        private Task<Industry> IndustryToEntityAsync(IndustryModel model)
        {
            return Task.FromResult(new Industry
            {
                Name = model.Name,
                Id = model.Id
            });
        }

        private async Task<Note> NoteToEntityAsync(NoteModel model, string requestId = "")
        {
            var note = Note.Empty;

            if (model.CreatedBy != null) note.CreatedBy = await _userProfileHelpers.UserProfileToEntityAsync(model.CreatedBy, requestId);
            if (model.CreatedDateTime == null)
            {
                note.CreatedDateTime = DateTimeOffset.Now;
            }
            else
            {
                note.CreatedDateTime = model.CreatedDateTime;
            }

            note.Id = model.Id ?? new Guid().ToString();
            note.NoteBody = model.NoteBody ?? String.Empty;

            return note;
        }

        private async Task<TeamMember> TeamMemberToEntityAsync(TeamMemberModel model, string requestId = "")
        {
            var teamMember = TeamMember.Empty;
            teamMember.Id = model.Id;
            teamMember.DisplayName = model.DisplayName;
            teamMember.Status = ActionStatus.FromValue(model.Status.Value);
            teamMember.AssignedRole = await _userProfileHelpers.RoleModelToEntityAsync(model.AssignedRole, requestId);
            teamMember.Fields = TeamMemberFields.Empty;
            teamMember.Fields.Mail = model.Mail;
            teamMember.Fields.Title = model.Title;
            teamMember.Fields.UserPrincipalName = model.UserPrincipalName;

            return teamMember;
        }

        private Task<Region> RegionToEntityAsync(RegionModel model)
        {
            return Task.FromResult(new Region
            {
                Name = model.Name,
                Id = model.Id
            });
        }
        #endregion
    }
}
