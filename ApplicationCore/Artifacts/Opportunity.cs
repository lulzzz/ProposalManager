// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using Newtonsoft.Json;
using ApplicationCore.Entities;
using ApplicationCore.Helpers;
using ApplicationCore.Serialization;

namespace ApplicationCore.Artifacts
{
    public class Opportunity : BaseArtifact<Opportunity>
    {
        public Opportunity()
        {
            ContentType = ContentType.Opportunity;
            Version = "1.0";
        }

        /// <summary>
        /// Content type of the opportunity
        /// </summary>
        [JsonProperty("contentType")]
        public new ContentType ContentType { get; private set; }

        /// <summary>
        /// Metadata of the opportunity
        /// </summary>
        [JsonProperty("metadata")]
        public new OpportunityMetadata Metadata { get; set; }

        /// <summary>
        /// Content of the opportunity
        /// </summary> // may have document, workflow etc
        [JsonProperty("content")]
        public new OpportunityContent Content { get; set; }

        /// <summary>
        /// Artifacts bag
        /// </summary>
        [JsonProperty("documentAttachments")]
        public IList<DocumentAttachment> DocumentAttachments { get; set; }

        /// <summary>
        /// Represents the empty opportunity. This field is read-only.
        /// </summary>
        public static Opportunity Empty
        {
            get => new Opportunity
            {
                Id = String.Empty,
                DisplayName = String.Empty,
                Reference = String.Empty,
                ContentType = ContentType.Opportunity,
                Version = "1.0",
                Metadata = OpportunityMetadata.Empty,
                Content = OpportunityContent.Empty,
                DocumentAttachments = new List<DocumentAttachment>()
            };
        }   
    }

    public class OpportunityMetadata
    {
        [JsonConverter(typeof(OpportunityStateConverter))]
        [JsonProperty("opportunityState")]
        public OpportunityState OpportunityState { get; set; }

        [JsonProperty("customer")]
        public Customer Customer { get; set; }

        [JsonProperty("dealSize")]
        public double DealSize { get; set; }

        [JsonProperty("annualRevenue")]
        public double AnnualRevenue { get; set; }

        [JsonProperty("openedDate")]
        public DateTimeOffset OpenedDate { get; set; }

        [JsonProperty("industry")]
        public Industry Industry { get; set; }

        [JsonProperty("region")]
        public Region Region { get; set; }

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
        public String Guarantees { get; set; }

        [JsonProperty("riskRating")]
        public int RiskRating { get; set; }

        [JsonProperty("opportunityChannelId")]
        public String OpportunityChannelId { get; set; }

        /// <summary>
        /// Represents the empty opportunity. This field is read-only.
        /// </summary>
        public static OpportunityMetadata Empty
        {
            get => new OpportunityMetadata
            {
                OpportunityState = OpportunityState.NoneEmpty,
                Customer = Customer.Empty,
                DealSize = 0.0,
                AnnualRevenue = 0.0,
                OpenedDate = DateTimeOffset.MinValue,
                Industry = Industry.Empty,
                Region = Region.Empty,
                Margin = 0.0,
                Rate = 0.0,
                DebtRatio = 0.0,
                Purpose = String.Empty,
                DisbursementSchedule = String.Empty,
                CollateralAmount = 0.0,
                Guarantees = String.Empty,
                RiskRating = 0,
                OpportunityChannelId = String.Empty
            };
        }    
    }

    public class OpportunityContent
    {
        [JsonProperty("teamMembers")]
        public IList<TeamMember> TeamMembers { get; set; }

        [JsonProperty("notes")]
        public IList<Note> Notes { get; set; }

        [JsonProperty("checklists")]
        public IList<Checklist> Checklists { get; set; }

        [JsonProperty("proposalDocument")]
        public ProposalDocument ProposalDocument { get; set; }

        [JsonProperty("customerDecision")]
        public CustomerDecision CustomerDecision { get; set; }

        /// <summary>
        /// Represents the empty opportunity. This field is read-only.
        /// </summary>
        public static OpportunityContent Empty
        {
            get => new OpportunityContent
            {
                TeamMembers = new List<TeamMember>(),
                Notes = new List<Note>(),
                Checklists = new List<Checklist>(),
                ProposalDocument = ProposalDocument.Empty
            };
        }  
    }

    public class OpportunityState : SmartEnum<OpportunityState, int>
    {
        public static OpportunityState NoneEmpty = new OpportunityState(nameof(NoneEmpty), 0);
        public static OpportunityState Creating = new OpportunityState(nameof(Creating), 1);
        public static OpportunityState InProgress = new OpportunityState(nameof(InProgress), 2);
        public static OpportunityState Assigned = new OpportunityState(nameof(Assigned), 3);
        public static OpportunityState Draft = new OpportunityState(nameof(Draft), 4);
        public static OpportunityState NotStarted = new OpportunityState(nameof(NotStarted), 5);
        public static OpportunityState InReview = new OpportunityState(nameof(InReview), 6);
        public static OpportunityState Blocked = new OpportunityState(nameof(Blocked), 7);
        public static OpportunityState Completed = new OpportunityState(nameof(Completed), 8);
        public static OpportunityState Submitted = new OpportunityState(nameof(Submitted), 9);
        public static OpportunityState Accepted = new OpportunityState(nameof(Accepted), 10);

        [JsonConstructor]
        protected OpportunityState(string name, int value) : base(name, value)
        {
        }
    }
}
