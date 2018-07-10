// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using SmartLink.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public interface IDestinationService
    {
        Task<DestinationPoint> AddDestinationPoint(string fileName, string documentId, DestinationPoint destinationPoint);
        Task<DestinationCatalog> GetDestinationCatalog(string fileName, string documentId);
        Task<IEnumerable<DestinationPoint>> GetDestinationPointBySourcePoint(Guid sourcePointId);
        Task DeleteDestinationPoint(Guid destinationPointId);
        Task DeleteSelectedDestinationPoint(IEnumerable<Guid> seletedDestinationPointIds);

        Task<IEnumerable<CustomFormat>> GetCustomFormats();
        Task<DestinationPoint> UpdateDestinationPointCustomFormat(DestinationPoint destinationPoint);
    }
}
