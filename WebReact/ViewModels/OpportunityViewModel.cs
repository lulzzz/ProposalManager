// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ApplicationCore.Artifacts;
using ApplicationCore.Helpers;
using ApplicationCore.Serialization;
using Newtonsoft.Json;
using WebReact.Models;
using WebReact.Serialization;

namespace WebReact.ViewModels
{
    public class OpportunityViewModel
    {
        public OpportunityViewModel()
        {
            Id = String.Empty;
            Reference = String.Empty;
            DisplayName = String.Empty;
            OpportunityState = OpportunityStateModel.NoneEmpty;
        }

        /// <summary>
        /// Unique identifier of the artifact
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Unique identifier of the artifact
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        /// <summary>
        /// Unique identifier of the artifact
        /// </summary>
        [JsonProperty("reference")]
        public string Reference { get; set; }

        [JsonProperty("version")]
        public string Version { get; set; }

        // ContentType not needed
        // Version not needed
        // Uri not needed
        // TypeName not needed

        // Metadata
        [JsonConverter(typeof(OpportunityStateModelConverter))]
        [JsonProperty("opportunityState")]
        public OpportunityStateModel OpportunityState { get; set; }

        [JsonProperty("customer")]
        public CustomerModel Customer { get; set; }

        [JsonProperty("dealSize")]
        public double DealSize { get; set; }

        [JsonProperty("annualRevenue")]
        public double AnnualRevenue { get; set; }

        [JsonProperty("openedDate")]
        public DateTimeOffset OpenedDate { get; set; }

        [JsonProperty("industry")]
        public IndustryModel Industry { get; set; }

        [JsonProperty("region")]
        public RegionModel Region { get; set; }

        [JsonProperty("margin")]
        public double Margin { get; set; }

        [JsonProperty("rate")]
        public double Rate { get; set; }

        [JsonProperty("debtRatio")]
        public double DebtRatio { get; set; }

        [JsonProperty("purpose")]
        public string Purpose { get; set; }

        [JsonProperty("disbursementSchedule")]
        public string DisbursementSchedule { get; set; }

        [JsonProperty("collateralAmount")]
        public Double CollateralAmount { get; set; }

        [JsonProperty("guarantees")]
        public string Guarantees { get; set; }

        [JsonProperty("riskRating")]
        public int RiskRating { get; set; }

        [JsonProperty("opportunityChannelId")]
        public String OpportunityChannelId { get; set; }

        // Content
        [JsonProperty("teamMembers")]
        public IList<TeamMemberModel> TeamMembers { get; set; }

        [JsonProperty("notes")]
        public IList<NoteModel> Notes { get; set; }

        [JsonProperty("checklists")]
        public IList<ChecklistModel> Checklists { get; set; }

        [JsonProperty("proposalDocument")]
        public ProposalDocumentModel ProposalDocument { get; set; }

        [JsonProperty("customerDecision")]
        public CustomerDecisionModel CustomerDecision { get; set; }

        // DocumentAttachments
        [JsonProperty("documentAttachments")]
        public IList<DocumentAttachmentModel> DocumentAttachments { get; set; }
    }

    public class OpportunityStateModel : SmartEnum<OpportunityStateModel, int>
    {
        public static OpportunityStateModel NoneEmpty = new OpportunityStateModel(nameof(NoneEmpty), 0);
        public static OpportunityStateModel Creating = new OpportunityStateModel(nameof(Creating), 1);
        public static OpportunityStateModel InProgress = new OpportunityStateModel(nameof(InProgress), 2);
        public static OpportunityStateModel Assigned = new OpportunityStateModel(nameof(Assigned), 3);
        public static OpportunityStateModel Draft = new OpportunityStateModel(nameof(Draft), 4);
        public static OpportunityStateModel NotStarted = new OpportunityStateModel(nameof(NotStarted), 5);
        public static OpportunityStateModel InReview = new OpportunityStateModel(nameof(InReview), 6);
        public static OpportunityStateModel Blocked = new OpportunityStateModel(nameof(Blocked), 7);
        public static OpportunityStateModel Completed = new OpportunityStateModel(nameof(Completed), 8);
        public static OpportunityStateModel Submitted = new OpportunityStateModel(nameof(Submitted), 9);
        public static OpportunityStateModel Accepted = new OpportunityStateModel(nameof(Accepted), 10);

        [JsonConstructor]
        protected OpportunityStateModel(string name, int value) : base(name, value)
        {
        }
    }
}

