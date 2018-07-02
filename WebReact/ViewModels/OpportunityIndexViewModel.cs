// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ApplicationCore;
using ApplicationCore.Artifacts;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using WebReact.Models;

namespace WebReact.ViewModels
{
    public class OpportunityIndexViewModel : BaseListViewModel<OpportunityIndexModel>
    {
        public OpportunityIndexViewModel() : base()
        {
        }
    }
}
