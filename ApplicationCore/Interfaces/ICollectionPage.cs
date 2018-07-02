// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;

namespace ApplicationCore.Interfaces
{
    /// <summary>
    /// Interface for collection pages.
    /// </summary>
    /// <typeparam name="T">The type of the collection.</typeparam>
    public interface ICollectionPage<T> : IList<T>
    {
        /// <summary>
        /// Skip token for this page, empty of there are no next results
        /// </summary>
        string SkipToken { get; }

        /// <summary>
        /// Number of items per page
        /// </summary>
        int ItemsPage { get; }

        /// <summary>
        /// Current page index
        /// </summary>
        int PageIndex { get; }

        /// <summary>
        /// The current page of the collection.
        /// </summary>
        IList<T> CurrentPage { get; }
    }
}
