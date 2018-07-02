// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;
using ApplicationCore.Artifacts;
using ApplicationCore.Serialization;


namespace ApplicationCore.Interfaces
{
    /// <summary>
    /// The interface IOpportunityCollectionPage.
    /// </summary>
    //[JsonConverter(typeof(InterfaceConverter<OpportunityCollectionPage>))]
    public interface IOpportunityCollectionPage : ICollectionPage<Opportunity>
    {
        // TODO: Finish implementation of interface

        /// <summary>
        /// Gets the next page <see cref="IOpportunityCollectionRequest"/> instance.
        /// </summary>
        //IOpportunityCollectionRequest NextPageRequest { get; }

        /// <summary>
        /// Initializes the NextPageRequest property.
        /// </summary>
        //void InitializeNextPageRequest(string nextPageLinkString);
    }
}
