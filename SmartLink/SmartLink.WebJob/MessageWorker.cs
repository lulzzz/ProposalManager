// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using Microsoft.Azure.WebJobs;
using Microsoft.WindowsAzure.Storage.Table;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure;
using Microsoft.Azure;
using System.Security.Cryptography;
using System.Diagnostics;
using Newtonsoft.Json.Linq;
using SmartLink.Service;
using SmartLink.Common;
using SmartLink.Entity;

namespace Smartlink.WebJob
{
    public class MessageWorker
    {
        private readonly ILogService _logService;
        private readonly IAzureStorageService _azureStorageService;
        private readonly ISourceService _sourceService;
        private readonly IDestinationService _destinationService;
        private readonly IDocumentService _documentService;

        private string _storageAccount = string.Empty;
        public MessageWorker(ILogService logService, ISourceService sourceService, IDestinationService destinationService, IAzureStorageService azureStorageService, IDocumentService documentService)
        {
            _logService = logService;
            _azureStorageService = azureStorageService;
            _sourceService = sourceService;
            _destinationService = destinationService;
            _documentService = documentService;
        }

        public void ProcessQueueMessage(
            [QueueTrigger(Constant.PUBLISH_QUEUE_NAME)] PublishedMessage message,
            [Table(Constant.PUBLISH_TABLE_NAME)] CloudTable tableBinding,
            TextWriter log)
        {
            try
            {
                var retValue = new PublishStatusEntity(message.PublishBatchId.ToString(), message.SourcePointId.ToString(), message.PublishHistoryId.ToString());

                var publishHistory = _sourceService.GetPublishHistoryById(message.PublishHistoryId).Result;

                if (publishHistory != null)
                {
                    var destinationPoints = _destinationService.GetDestinationPointBySourcePoint(publishHistory.SourcePointId);
                    var groupedDestinationPoints = destinationPoints.Result.GroupBy(o => o.CatalogId);
                    var publishValue = publishHistory.Value;
                    IDictionary<string, Task<DocumentUpdateResult>> tasks = new Dictionary<string, Task<DocumentUpdateResult>>();
                    foreach (var sameCatalogDestinationPoints in groupedDestinationPoints)
                    {
                        try
                        {
                            var documentId = sameCatalogDestinationPoints.First().Catalog.DocumentId;
                            var points = sameCatalogDestinationPoints.Select(o => o);
                            tasks.Add(documentId, _documentService.UpdateBookmrkValue(documentId, points, publishValue));
                        }
                        catch (Exception ex)
                        {
                            log.Write($"Publish the source point to file '{message.SourcePointId}' failed due to {ex.ToString()}");
                        }
                    }
                    Task.WaitAll(tasks.Values.ToArray());
                    var errorItems = tasks.Where(o => o.Value.Result.IsSuccess == false || o.Value.IsFaulted);
                    retValue.Comments = String.Join("\n\n", tasks.Select(o => $"{o.Key}:\t{String.Join("\n",o.Value.Result.Message)}"));
                    if (errorItems.Count() > 0)
                    {
                        retValue.Status = PublishStatus.Error;
                        retValue.ErrorSummary = $"Update files: {String.Join(";", errorItems.Select(o => o.Key))} failed";
                        retValue.ErrorDetail = String.Join("\n", errorItems.SelectMany(o => o.Value.Result.Message));
                        log.Write($"Update the documents {retValue.ErrorSummary} failed due to {retValue.ErrorSummary} ");
                    }
                    else
                    {
                        retValue.Status = PublishStatus.Completed;
                        log.Write($"Update the documents successfully.");
                    }
                }
                else
                {
                    retValue.Status = PublishStatus.Error;
                    retValue.ErrorSummary = "The publish history cannot be found.";
                    log.Write($"The publish history related to the source point: '{message.SourcePointId}' cannot be found.");
                }

                tableBinding.Execute(TableOperation.InsertOrReplace(retValue));
                log.Write("Publish is finished.");
            }
            catch(Exception ex)
            {
                log.Write($"Publish the source point: '{message.SourcePointId}' failed due to {ex.ToString()}");
            }
        }
    }
}
