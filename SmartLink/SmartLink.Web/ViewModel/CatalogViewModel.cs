// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SmartLink.Web.ViewModel
{
    public class CatalogViewModel
    {
        public string Name { get; set; }

        public string DocumentId { get; set; }

        public bool IsDeleted { get; set; }
    }
}