// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SmartLink.Entity;
using AutoMapper;
using System.Data.Entity;
using Newtonsoft.Json;
using SmartLink.Common;
using System.Threading;
using Microsoft.WindowsAzure.Storage.Table;

namespace SmartLink.Service
{
    public class SourceService : ISourceService
    {
        protected readonly SmartlinkDbContext _dbContext;
        protected readonly IMapper _mapper;
        protected readonly IAzureStorageService _azureStorageService;
        protected readonly ILogService _logService;
        protected readonly IUserProfileService _userProfileService;

        public SourceService(SmartlinkDbContext dbContext, IMapper mapper, IAzureStorageService azureStorageService, ILogService logService, IUserProfileService userProfileService)
        {
            _dbContext = dbContext;
            _mapper = mapper;
            _azureStorageService = azureStorageService;
            _logService = logService;
            _userProfileService = userProfileService;
        }

        public async Task<SourcePoint> AddSourcePoint(string fileName, string documentId, SourcePoint sourcePoint)
        {
            try
            {
                var sourceCatalog = _dbContext.SourceCatalogs.FirstOrDefault(o => o.DocumentId == documentId);
                bool addSourceCatalog = (sourceCatalog == null);
                if (addSourceCatalog)
                {
                    try
                    {
                        sourceCatalog = new SourceCatalog() { Name = fileName, DocumentId = documentId };
                        _dbContext.SourceCatalogs.Add(sourceCatalog);
                    }
                    catch (Exception ex)
                    {
                        var entity = new LogEntity()
                        {
                            LogId = "30006",
                            Action = Constant.ACTIONTYPE_ADD,
                            ActionType = ActionTypeEnum.ErrorLog,
                            PointType = Constant.POINTTYPE_SOURCECATALOG,
                            Message = ".Net Error",
                        };
                        entity.Subject = $"{entity.LogId} - {entity.Action} - {entity.PointType} - Error";
                        await _logService.WriteLog(entity);

                        throw new ApplicationException("Add Source Catalog failed", ex);
                    }
                }

                sourcePoint.Created = DateTime.Now.ToUniversalTime().ToPSTDateTime();
                sourcePoint.Creator = _userProfileService.GetCurrentUser().Username;

                sourceCatalog.SourcePoints.Add(sourcePoint);

                _dbContext.PublishedHistories.Add(new PublishedHistory()
                {
                    Name = sourcePoint.Name,
                    Position = sourcePoint.Position,
                    Value = sourcePoint.Value,
                    PublishedDate = sourcePoint.Created,
                    PublishedUser = sourcePoint.Creator,
                    SourcePointId = sourcePoint.Id
                });

                await _dbContext.SaveChangesAsync();

                if (addSourceCatalog)
                {
                    await _logService.WriteLog(new LogEntity()
                    {
                        LogId = "30003",
                        Action = Constant.ACTIONTYPE_ADD,
                        PointType = Constant.POINTTYPE_SOURCECATALOG,
                        ActionType = ActionTypeEnum.AuditLog,
                        Message = $"Add Source Catalog named {sourceCatalog.Name}."
                    });
                }
                await _logService.WriteLog(new LogEntity()
                {
                    LogId = "10001",
                    Action = Constant.ACTIONTYPE_ADD,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    ActionType = ActionTypeEnum.AuditLog,
                    Message = $"Create source point named {sourcePoint.Name} in the location: {sourcePoint.Position}, value: {sourcePoint.Value} in the excel file named: {sourceCatalog.FileName} by {sourcePoint.Creator}"
                });

            }
            catch (ApplicationException ex)
            {
                throw ex.InnerException;
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "10005",
                    Action = Constant.ACTIONTYPE_ADD,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLog(logEntity);
                throw ex;
            }
            return sourcePoint;
        }

        public async Task<SourcePoint> EditSourcePoint(SourcePoint sourcePoint)
        {
            try
            {
                var previousSourcePoint = await _dbContext.SourcePoints.Include(o => o.Catalog).FirstOrDefaultAsync(o => o.Id == sourcePoint.Id);

                if (previousSourcePoint != null)
                {
                    previousSourcePoint.Name = sourcePoint.Name;
                    previousSourcePoint.Position = sourcePoint.Position;
                    previousSourcePoint.RangeId = sourcePoint.RangeId;
                    previousSourcePoint.Value = sourcePoint.Value;
                    previousSourcePoint.NamePosition = sourcePoint.NamePosition;
                    previousSourcePoint.NameRangeId = sourcePoint.NameRangeId;
                    previousSourcePoint.PublishedHistories = (await _dbContext.PublishedHistories.Where(o => o.SourcePointId == previousSourcePoint.Id).ToArrayAsync()).OrderByDescending(o => o.PublishedDate).ToArray();
                }

                await _dbContext.SaveChangesAsync();
                await _logService.WriteLog(new LogEntity()
                {
                    LogId = "10002",
                    Action = Constant.ACTIONTYPE_EDIT,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    ActionType = ActionTypeEnum.AuditLog,
                    Message = $"Edit source point by {_userProfileService.GetCurrentUser().Username} Previous value: source point named: {previousSourcePoint.Name} in the location: {previousSourcePoint.Position} value: {previousSourcePoint.Value} in the excel file named: {previousSourcePoint.Catalog.FileName} " +
                              $"Current value: source point named: {sourcePoint.Name} in the location {sourcePoint.Position} value: {sourcePoint.Value} in the excel file name: {previousSourcePoint.Catalog.FileName}"
                });

                return previousSourcePoint;
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "10008",
                    Action = Constant.ACTIONTYPE_EDIT,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLog(logEntity);
                throw ex;
            }
        }
        public async Task<int> DeleteSourcePoint(Guid sourcePointId)
        {
            try
            {
                var sourcePoint = _dbContext.SourcePoints.Include(o => o.Catalog).FirstOrDefault(o => o.Id == sourcePointId);
                if (sourcePoint == null)
                {
                    throw new NullReferenceException(string.Format("Sourcepoint: {0} is not existed", sourcePointId));
                }
                if (sourcePoint.Status == SourcePointStatus.Deleted)
                {
                    return await Task.FromResult<int>(0);
                }
                else
                {
                    sourcePoint.Status = SourcePointStatus.Deleted;
                }
                var task = await _dbContext.SaveChangesAsync();
                await _logService.WriteLog(new LogEntity()
                {
                    LogId = "10003",
                    Action = Constant.ACTIONTYPE_DELETE,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    ActionType = ActionTypeEnum.AuditLog,
                    Message = $"Delete source point named: {sourcePoint.Name} in the excel file named: {sourcePoint.Catalog.FileName}"
                });
                return task;
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "10006",
                    Action = Constant.ACTIONTYPE_DELETE,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLog(logEntity);
                throw ex;
            }

        }

        public async Task DeleteSelectedSourcePoint(IEnumerable<Guid> selectedSourcePointIds)
        {
            try
            {
                foreach (var sourcePointId in selectedSourcePointIds)
                {
                    var sourcePoint = _dbContext.SourcePoints.Include(o => o.Catalog).FirstOrDefault(o => o.Id == sourcePointId);
                    if (sourcePoint == null)
                    {
                        throw new NullReferenceException(string.Format("Sourcepoint: {0} is not existed", sourcePointId));
                    }
                    if (sourcePoint.Status != SourcePointStatus.Deleted)
                    {
                        sourcePoint.Status = SourcePointStatus.Deleted;
                    }
                    var task = await _dbContext.SaveChangesAsync();
                    await _logService.WriteLog(new LogEntity()
                    {
                        LogId = "10003",
                        Action = Constant.ACTIONTYPE_DELETE,
                        PointType = Constant.POINTTYPE_SOURCEPOINT,
                        ActionType = ActionTypeEnum.AuditLog,
                        Message = $"Delete source point named: {sourcePoint.Name} in the excel file named: {sourcePoint.Catalog.FileName}"
                    });
                }
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "10006",
                    Action = Constant.ACTIONTYPE_DELETE,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLog(logEntity);
                throw ex;
            }

        }

        public async Task<SourceCatalog> GetSourceCatalog(string fileName, string documentId)
        {
            try
            {
                var sourceCatalog = await _dbContext.SourceCatalogs.Where(o => o.DocumentId == documentId).FirstOrDefaultAsync();
                if (sourceCatalog != null)
                {
                    var sourcePointArray = (await _dbContext.SourcePoints.Where(o => o.Status == SourcePointStatus.Created && o.CatalogId == sourceCatalog.Id)
                         .Include(o => o.DestinationPoints).ToArrayAsync())
                         .OrderByDescending(o => o.Name).ToArray();
                    var sourcePointIds = sourcePointArray.Select(point => point.Id).ToArray();
                    var publishedHistories = await (from pb in _dbContext.PublishedHistories
                                                    where sourcePointIds.Contains(pb.SourcePointId)
                                                    select pb).ToArrayAsync();
                    foreach (var item in sourcePointArray)
                    {
                        item.PublishedHistories = publishedHistories.Where(pb => pb.SourcePointId == item.Id)
                                                                    .OrderByDescending(pb => pb.PublishedDate).ToArray();
                    }
                    sourceCatalog.SourcePoints = sourcePointArray;

                    if (!sourceCatalog.Name.Equals(fileName))
                    {
                        sourceCatalog.Name = fileName;
                        await _dbContext.SaveChangesAsync();
                    }

                    await _logService.WriteLog(new LogEntity()
                    {
                        LogId = "30001",
                        Action = Constant.ACTIONTYPE_GET,
                        PointType = Constant.POINTTYPE_SOURCECATALOG,
                        ActionType = ActionTypeEnum.AuditLog,
                        Message = $"Get source catalog named {sourceCatalog.Name}"
                    });
                }
                return sourceCatalog;
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "30004",
                    Action = Constant.ACTIONTYPE_GET,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCECATALOG,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLog(logEntity);
                throw ex;
            }
        }

        public async Task<SourceCatalog> GetSourceCatalog(string documentId)
        {
            try
            {
                var sourceCatalog = await _dbContext.SourceCatalogs.Where(o => o.DocumentId == documentId).FirstOrDefaultAsync();
                if (sourceCatalog != null)
                {
                    var sourcePointArray = (await _dbContext.SourcePoints.Where(o => o.Status == SourcePointStatus.Created && o.CatalogId == sourceCatalog.Id)
                         .Include(o => o.DestinationPoints).ToArrayAsync())
                         .OrderByDescending(o => o.Name).ToArray();
                    var sourcePointIds = sourcePointArray.Select(point => point.Id).ToArray();
                    var publishedHistories = await (from pb in _dbContext.PublishedHistories
                                                    where sourcePointIds.Contains(pb.SourcePointId)
                                                    select pb).ToArrayAsync();
                    foreach (var item in sourcePointArray)
                    {
                        item.PublishedHistories = publishedHistories.Where(pb => pb.SourcePointId == item.Id)
                                                                    .OrderByDescending(pb => pb.PublishedDate).ToArray();
                    }
                    sourceCatalog.SourcePoints = sourcePointArray;

                    await _logService.WriteLog(new LogEntity()
                    {
                        LogId = "30001",
                        Action = Constant.ACTIONTYPE_GET,
                        PointType = Constant.POINTTYPE_SOURCECATALOG,
                        ActionType = ActionTypeEnum.AuditLog,
                        Message = $"Get source catalog named {sourceCatalog.Name}"
                    });
                }
                return sourceCatalog;
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "30004",
                    Action = Constant.ACTIONTYPE_GET,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCECATALOG,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLog(logEntity);
                throw ex;
            }
        }

        public async Task<PublishSourcePointResult> PublishSourcePointList(IEnumerable<PublishSourcePointForm> publishSourcePointForms)
        {
            try
            {
                var sourcePointIdList = publishSourcePointForms.Select(o => o.SourcePointId).ToArray();
                var sourcePointList = _dbContext.SourcePoints.Include(o => o.Catalog).Where(o => sourcePointIdList.Contains(o.Id)).ToList();
                var currentUser = _userProfileService.GetCurrentUser();

                //Update database
                IList<PublishedHistory> histories = new List<PublishedHistory>();
                foreach (var sourcePoint in sourcePointList)
                {
                    var sourcePointForm = publishSourcePointForms.First(o => o.SourcePointId == sourcePoint.Id);
                    sourcePoint.Value = sourcePointForm.CurrentValue;
                    sourcePoint.Position = sourcePointForm.Position;
                    sourcePoint.Name = sourcePointForm.Name;
                    sourcePoint.NamePosition = sourcePointForm.NamePosition;

                    var history = new PublishedHistory()
                    {
                        Name = sourcePoint.Name,
                        Position = sourcePoint.Position,
                        Value = sourcePoint.Value,
                        PublishedDate = DateTime.Now.ToUniversalTime().ToPSTDateTime(),
                        PublishedUser = currentUser.Username,
                        SourcePointId = sourcePoint.Id
                    };

                    _dbContext.PublishedHistories.Add(history);

                    histories.Add(history);
                }
                await _dbContext.SaveChangesAsync();

                //Update Table Storage
                var table = _azureStorageService.GetTable(Constant.PUBLISH_TABLE_NAME);

                var batchId = Guid.NewGuid();

                //A single batch operation can include up to 100 entities, so seperate histories into several batch when histories 
                //https://docs.microsoft.com/en-us/azure/storage/storage-dotnet-how-to-use-tables
                var batchCount = Math.Ceiling((float)histories.Count() / Constant.AZURETABLE_BATCH_COUNT);
                var batchTasks = new List<Task<IList<TableResult>>>();
                for (int i = 0; i < batchCount; i++)
                {
                    var historiesPerBatch = histories.Skip(Constant.AZURETABLE_BATCH_COUNT * i).Take(Constant.AZURETABLE_BATCH_COUNT);
                    var batchOpt = new TableBatchOperation();
                    foreach (var item in historiesPerBatch)
                    {
                        batchOpt.Insert(new PublishStatusEntity(batchId.ToString(), item.SourcePointId.ToString(), item.Id.ToString()));
                    }
                    batchTasks.Add(table.ExecuteBatchAsync(batchOpt));
                }

                Task.WaitAll(batchTasks.ToArray());
                IList<SourcePoint> errorSourcePoints = new List<SourcePoint>();
                var batchResults = batchTasks.SelectMany(o => o.Result).ToArray();
                for (int i = 0; i < batchResults.Count(); i++)
                {
                    if (!(batchResults[i].HttpStatusCode == 200 || batchResults[i].HttpStatusCode == 204))
                    {
                        errorSourcePoints.Add(sourcePointList[i]);
                    }
                }
                if (errorSourcePoints.Count() > 0)
                {
                    throw new SourcePointException(sourcePointList, string.Concat(batchResults.Where(o => !(o.HttpStatusCode == 200 || o.HttpStatusCode == 204)).Select(o => o.Result)), null);
                }

                //Push message to queue
                IList<Task> writeQueueTask = new List<Task>();
                foreach (var history in histories)
                {
                    var message = new PublishedMessage() { PublishHistoryId = history.Id, SourcePointId = history.SourcePointId, PublishBatchId = batchId };
                    writeQueueTask.Add(Task.Run(async () =>
                       {
                           await _azureStorageService.WriteMessageToQueue(JsonConvert.SerializeObject(message), Constant.PUBLISH_QUEUE_NAME);
                           await _logService.WriteLog(new LogEntity()
                           {
                               LogId = "10004",
                               Action = Constant.ACTIONTYPE_PUBLISH,
                               ActionType = ActionTypeEnum.AuditLog,
                               PointType = Constant.POINTTYPE_SOURCEPOINT,
                               Message = $"Publish source point named: {history.Name} in the location {history.Position}, value: {history.Value} in the excel file named:{history.SourcePoint.Catalog.Name} by {currentUser.Username}"
                           });
                       }));
                }
                await Task.WhenAll(writeQueueTask);

                //var publishedHistories = await (from pb in _dbContext.PublishedHistories
                //                                where sourcePointIdList.Contains(pb.SourcePointId)
                //                                select pb).ToArrayAsync();
                foreach (var item in sourcePointList)
                {
                    item.PublishedHistories = histories.Where(h => h.SourcePointId == item.Id).ToArray();
                }

                return new PublishSourcePointResult() { BatchId = batchId, SourcePoints = sourcePointList };
            }
            catch (SourcePointException ex)
            {
                foreach (var sourcePoint in ex.ErrorSourcePoints)
                {
                    var logEntity = new LogEntity()
                    {
                        LogId = "10007",
                        Action = Constant.ACTIONTYPE_PUBLISH,
                        ActionType = ActionTypeEnum.ErrorLog,
                        PointType = Constant.POINTTYPE_SOURCEPOINT,
                        Message = ".Net Error",
                        Detail = $"{sourcePoint.Id}-{sourcePoint.Name}-{ex.ToString()}"
                    };
                    logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                    await _logService.WriteLog(logEntity);
                }

                throw ex;
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "10007",
                    Action = Constant.ACTIONTYPE_PUBLISH,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLog(logEntity);
                throw ex;
            }
        }

        public async Task<PublishedHistory> GetPublishHistoryById(Guid publishHistoryId)
        {
            var publishHistory = await _dbContext.PublishedHistories.Include(o => o.SourcePoint.Catalog).FirstOrDefaultAsync(o => o.Id == publishHistoryId);
            if (publishHistory != null)
            {
                publishHistory.SourcePoint.SerializeCatalog = true;
                publishHistory.SourcePoint.Catalog.SerializeSourcePoints = false;
            }
            return publishHistory;
        }

        public async Task<IEnumerable<DocumentCheckResult>> GetAllCatalogs()
        {
            var sourceCatalogs = await _dbContext.SourceCatalogs.Where(o => !string.IsNullOrEmpty(o.DocumentId)).ToListAsync();
            var destinationCatalogs = await _dbContext.DestinationCatalogs.Where(o => !string.IsNullOrEmpty(o.DocumentId)).ToListAsync();
            var catalog = new List<DocumentCheckResult>();
            catalog.AddRange(sourceCatalogs.Select(o => new DocumentCheckResult() { DocumentId = o.DocumentId, DocumentType = DocumentTypes.SourcePoint }).ToList());
            catalog.AddRange(destinationCatalogs.Select(o => new DocumentCheckResult() { DocumentId = o.DocumentId, DocumentType = DocumentTypes.DestinationPoint }).ToList());
            return catalog;
        }

        public async Task<IEnumerable<DocumentCheckResult>> UpdateDocumentUrlById(IEnumerable<DocumentCheckResult> documents)
        {
            foreach (var document in documents)
            {
                if (document.DocumentType == DocumentTypes.SourcePoint)
                {
                    var catalog = await _dbContext.SourceCatalogs.Where(o => o.DocumentId == document.DocumentId).FirstOrDefaultAsync();
                    if (catalog != null)
                    {
                        if (document.IsDeleted)
                        {
                            document.IsUpdated = true;
                            document.Message = string.Format("DocumentId: {0}, The file {1} has been deleted", document.DocumentId, catalog.Name);
                            catalog.IsDeleted = true;
                        }
                        else
                        {
                            if (!catalog.Name.Equals(document.DocumentUrl, StringComparison.CurrentCultureIgnoreCase))
                            {
                                document.IsUpdated = true;
                                document.Message = string.Format("DocumentId: {0}, Updated from {1} to {2}", document.DocumentId, catalog.Name, document.DocumentUrl);
                                catalog.Name = document.DocumentUrl;
                                catalog.IsDeleted = false;
                            }
                            else
                            {
                                if (catalog.IsDeleted)
                                {
                                    document.IsUpdated = true;
                                    catalog.IsDeleted = false;
                                }
                            }
                        }
                    }
                }
                else if (document.DocumentType == DocumentTypes.DestinationPoint)
                {
                    var catalog = await _dbContext.DestinationCatalogs.Where(o => o.DocumentId == document.DocumentId).FirstOrDefaultAsync();
                    if (catalog != null)
                    {
                        if (document.IsDeleted)
                        {
                            document.IsUpdated = true;
                            document.Message = string.Format("DocumentId: {0}, The file {1} has been deleted", document.DocumentId, catalog.Name);
                            catalog.IsDeleted = true;
                        }
                        else
                        {
                            if (!catalog.Name.Equals(document.DocumentUrl, StringComparison.CurrentCultureIgnoreCase))
                            {
                                document.IsUpdated = true;
                                document.Message = string.Format("DocumentId: {0}, Updated from {1} to {2}", document.DocumentId, catalog.Name, document.DocumentUrl);
                                catalog.Name = document.DocumentUrl;
                                catalog.IsDeleted = false;
                            }
                            else
                            {
                                if (catalog.IsDeleted)
                                {
                                    document.IsUpdated = true;
                                    catalog.IsDeleted = false;
                                }
                            }
                        }
                    }
                }
            }
            if (documents.Any(o => o.IsUpdated))
            {
                await _dbContext.SaveChangesAsync();
            }
            return documents;
        }

        public async Task<IEnumerable<CloneForm>> CheckCloneFileStatus(IEnumerable<CloneForm> files)
        {
            var destinationCatalogIds = new List<string>();
            foreach (var item in files)
            {
                if (item.IsExcel)
                {
                    var catalog = await _dbContext.SourceCatalogs.Where(o => o.DocumentId.Equals(item.DocumentId, StringComparison.CurrentCultureIgnoreCase)).FirstOrDefaultAsync();
                    if (catalog != null)
                    {
                        var sourcePoints = await _dbContext.SourcePoints.Where(o => o.Status == SourcePointStatus.Created && o.CatalogId == catalog.Id)
                        .Include(o => o.DestinationPoints)
                        .ToArrayAsync();
                        if (sourcePoints.Count() > 0)
                        {
                            item.Clone = true;
                            foreach (var sourcePoint in sourcePoints)
                            {
                                foreach (var destinationPoint in sourcePoint.DestinationPoints)
                                {
                                    var destinationCatalog = await _dbContext.DestinationCatalogs.Where(o => o.Id.Equals(destinationPoint.CatalogId)).FirstOrDefaultAsync();
                                    if (destinationCatalog != null)
                                    {
                                        destinationCatalogIds.Add(destinationCatalog.DocumentId);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            foreach (var item in files)
            {
                if (item.IsWord)
                {
                    if (destinationCatalogIds.Any(o => o.Equals(item.DocumentId, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        item.Clone = true;
                    }
                }
            }
            return files;
        }

        public async Task CloneFiles(IEnumerable<CloneForm> files)
        {
            try
            {
                var filesWillClone = files.Where(o => o.Clone);
                var clonedSourcePoints = new Dictionary<Guid, SourcePoint>();
                foreach (var item in filesWillClone)
                {
                    if (item.IsExcel)
                    {
                        var catalog = await _dbContext.SourceCatalogs.Where(o => o.DocumentId.Equals(item.DocumentId, StringComparison.CurrentCultureIgnoreCase)).FirstOrDefaultAsync();
                        if (catalog != null)
                        {
                            var sourcePoints = await _dbContext.SourcePoints.Where(o => o.Status == SourcePointStatus.Created && o.CatalogId == catalog.Id)
                                .Include(o => o.DestinationPoints)
                                .ToArrayAsync();

                            var sourcePointIds = sourcePoints.Select(point => point.Id).ToArray();
                            var publishedHistories = await (from pb in _dbContext.PublishedHistories
                                                            where sourcePointIds.Contains(pb.SourcePointId)
                                                            select pb).ToArrayAsync();

                            #region Add new source catalog

                            SourceCatalog newSourceCatalog = new SourceCatalog() { Name = item.DestinationFileUrl, DocumentId = item.DestinationFileDocumentId, IsDeleted = false };
                            _dbContext.SourceCatalogs.Add(newSourceCatalog);

                            #endregion

                            #region Add new source point

                            foreach (var sourcePoint in sourcePoints)
                            {
                                var lastPublishedValue = publishedHistories.Where(o => o.SourcePointId == sourcePoint.Id).OrderByDescending(o => o.PublishedDate).FirstOrDefault().Value;
                                var newSourcePoint = new SourcePoint()
                                {
                                    Name = sourcePoint.Name,
                                    RangeId = sourcePoint.RangeId,
                                    Position = sourcePoint.Position,
                                    Value = sourcePoint.Value,
                                    Created = DateTime.Now.ToUniversalTime().ToPSTDateTime(),
                                    Creator = _userProfileService.GetCurrentUser().Username,
                                    Status = SourcePointStatus.Created,
                                    NamePosition = sourcePoint.NamePosition,
                                    NameRangeId = sourcePoint.NameRangeId,
                                    SourceType = sourcePoint.SourceType
                                };

                                _dbContext.PublishedHistories.Add(new PublishedHistory()
                                {
                                    Name = sourcePoint.Name,
                                    Position = sourcePoint.Position,
                                    Value = "Cloned",
                                    PublishedUser = _userProfileService.GetCurrentUser().Username,
                                    PublishedDate = DateTime.Now.ToUniversalTime().ToPSTDateTime(),
                                    SourcePointId = newSourcePoint.Id
                                });

                                _dbContext.PublishedHistories.Add(new PublishedHistory()
                                {
                                    Name = sourcePoint.Name,
                                    Position = sourcePoint.Position,
                                    Value = lastPublishedValue,
                                    PublishedUser = _userProfileService.GetCurrentUser().Username,
                                    PublishedDate = DateTime.Now.AddSeconds(1).ToUniversalTime().ToPSTDateTime(),
                                    SourcePointId = newSourcePoint.Id
                                });

                                newSourceCatalog.SourcePoints.Add(newSourcePoint);

                                clonedSourcePoints.Add(sourcePoint.Id, newSourcePoint);
                            }

                            #endregion
                        }
                    }
                }

                foreach (var item in filesWillClone)
                {
                    if (item.IsWord)
                    {
                        var catalog = await _dbContext.DestinationCatalogs.Where(o => o.DocumentId.Equals(item.DocumentId, StringComparison.CurrentCultureIgnoreCase)).FirstOrDefaultAsync();
                        if (catalog != null)
                        {
                            var destinationPoints = await _dbContext.DestinationPoints.Where(o => o.CatalogId == catalog.Id)
                                .Include(o => o.CustomFormats)
                                .Include(o => o.ReferencedSourcePoint)
                                .Include(o => o.ReferencedSourcePoint.Catalog).ToArrayAsync();

                            var sourcePointDocumentIds = destinationPoints.Select(o => o.ReferencedSourcePoint.Catalog.DocumentId);
                            var clonedExcelDocumentIds = files.Where(o => o.Clone && o.IsExcel).Select(o => o.DocumentId);
                            if (sourcePointDocumentIds.Any(o => clonedExcelDocumentIds.Any(p => p.Equals(o, StringComparison.CurrentCultureIgnoreCase))))
                            {
                                #region Add new destination catalog

                                DestinationCatalog newDestinationCatalog = new DestinationCatalog() { Name = item.DestinationFileUrl, DocumentId = item.DestinationFileDocumentId, IsDeleted = false };
                                _dbContext.DestinationCatalogs.Add(newDestinationCatalog);

                                #endregion

                                #region Add new destination point

                                foreach (var destinationPoint in destinationPoints)
                                {
                                    var referencedSourcePoint = clonedSourcePoints.FirstOrDefault(o => o.Key == destinationPoint.ReferencedSourcePoint.Id).Value;
                                    if (referencedSourcePoint != null)
                                    {
                                        DestinationPoint newDestinationPoint = new DestinationPoint()
                                        {
                                            RangeId = destinationPoint.RangeId,
                                            Created = DateTime.Now.ToUniversalTime().ToPSTDateTime(),
                                            Creator = _userProfileService.GetCurrentUser().Username,
                                            DecimalPlace = destinationPoint.DecimalPlace,
                                            DestinationType = destinationPoint.DestinationType
                                        };

                                        var newCustomFormatIds = destinationPoint.CustomFormats.Select(o => o.Id);
                                        var newCustomFormats = _dbContext.CustomFormats.Where(o => newCustomFormatIds.Contains(o.Id));
                                        foreach (var format in newCustomFormats)
                                        {
                                            newDestinationPoint.CustomFormats.Add(format);
                                        }

                                        newDestinationCatalog.DestinationPoints.Add(newDestinationPoint);

                                        newDestinationPoint.ReferencedSourcePoint = referencedSourcePoint;
                                        _dbContext.DestinationPoints.Add(newDestinationPoint);
                                    }
                                }

                                #endregion
                            }
                        }
                    }
                }

                await _dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                var entity = new LogEntity()
                {
                    LogId = "50001",
                    Action = Constant.ACTIONTYPE_CLONE,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_CLONE,
                    Message = ".Net Error",
                };
                entity.Subject = $"{entity.LogId} - {entity.Action} - {entity.PointType} - Error";
                await _logService.WriteLog(entity);

                throw new ApplicationException("Clone folder failed", ex);
            }
        }
    }

    class SourcePointException : Exception
    {
        public IList<SourcePoint> ErrorSourcePoints { get; protected set; }
        public SourcePointException(IList<SourcePoint> errorSourcePoints, string message, Exception innerException) : base(message, innerException)
        {
            ErrorSourcePoints = errorSourcePoints;
        }
    }

}
