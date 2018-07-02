// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.WindowsAzure.Storage.Table;

namespace SmartLink.Service
{
    public interface IAzureStorageService
    {
        Task WriteMessageToQueue(string queueMessage, string queueName);
        CloudTable GetTable(string tableName);
    }
}
