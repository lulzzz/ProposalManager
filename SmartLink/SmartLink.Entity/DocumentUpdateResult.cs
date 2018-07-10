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
    public class DocumentUpdateResult
    {
        public bool IsSuccess { get; set; }
        public List<string> Message { get; set; }
        public DocumentUpdateResult()
        {
            Message = new List<string>();
        }
    }
}
