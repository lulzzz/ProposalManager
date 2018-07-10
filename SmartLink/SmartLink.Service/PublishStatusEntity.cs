// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Entity
{
    public class PublishStatusEntity : TableEntity
    {
        public PublishStatusEntity()
        {

        }
        public PublishStatusEntity(string publishBatchId, string sourcePointId, string publishHistoryId)
        {
            this.PartitionKey = publishBatchId;
            this.RowKey = publishHistoryId;
            SourcePointId = sourcePointId;
            PublishHistoryId = publishHistoryId;
            Status = PublishStatus.InProgess;
        }
        public string PublishBatchId
        {
            get
            {
                return PartitionKey;
            }
        }
        public string PublishHistoryId { get; set; }
        public string SourcePointId { get; set; }

        public string StatusValue { get; set; }
        [IgnoreProperty]
        public PublishStatus Status
        {
            get { return (PublishStatus)Enum.Parse(typeof(PublishStatus),StatusValue); }
            set { StatusValue = value.ToString(); }
        }

        public string ErrorSummary { get; set; }
        public string ErrorDetail { get; set; }
        public string Comments { get; set; }
    }
}
