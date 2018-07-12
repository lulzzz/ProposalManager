// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using ApplicationCore.Helpers;
using ApplicationCore.Interfaces;

namespace ApplicationCore.Artifacts
{
    public class OpportunityCollectionPage : CollectionPage<Opportunity>, IOpportunityCollectionPage
    {

        public OpportunityCollectionPage() : base()
        {
        }

        public OpportunityCollectionPage(IList<Opportunity> currentPage) : base(currentPage)
        {
        }


    }
}
