// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Text;

namespace ApplicationCore.Helpers
{
    /// <summary>
    /// A query option to be added to the query parameters list.
    /// </summary>
    public class QueryParam : Option
    {
        /// <summary>
        /// Create a query option.
        /// </summary>
        /// <param name="name">The name of the query option, or parameter.</param>
        /// <param name="value">The value of the query option.</param>
        public QueryParam(string name, string value)
            : base(name, value)
        {
        }
    }
}
