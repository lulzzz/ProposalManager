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
    public enum PublishStatus
    {
        InProgess,
        Completed,
        Error
    }
    public class PublishStatusItem
    {
        public string PublishBatchId { get; set; }
        public string SourcePointId { get; set; }
        public string Status { get; set; }
        public string ErrorSummary { get; set; }
        public string ErrorDetail { get; set; }
    }
}
