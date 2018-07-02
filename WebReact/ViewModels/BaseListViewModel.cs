// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Text;
using ApplicationCore;
using ApplicationCore.Helpers;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace WebReact.ViewModels
{
    /// <summary>
    /// Base abstract class for ListViewMoels
    /// </summary>
    public abstract class BaseListViewModel<T>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public BaseListViewModel()
        {
            ItemsList = new List<T>();
            PaginationInfo = new PaginationInfoViewModel
            {
                TotalItems = 0,
                ItemsPerPage = 0,
                ActualPage = 0,
                TotalPages = 0,
                Previous = String.Empty,
                Next = String.Empty
            };
        }

        public IList<T> ItemsList { get; set; }

        public PaginationInfoViewModel PaginationInfo { get; set; }
    }
}
