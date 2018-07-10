// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Entity
{
    public class CloneForm
    {
        public string Name { get; set; }
        public string Url { get; set; }
        public string DocumentId { get; set; }
        public bool IsExcel { get; set; }
        public bool IsWord { get; set; }
        public bool Clone { get; set; }
        public string DestinationFileUrl { get; set; }
        public string DestinationFileDocumentId { get; set; }
    }
}
