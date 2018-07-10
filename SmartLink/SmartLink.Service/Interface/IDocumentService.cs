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
    public interface IDocumentService
    {
        Task<DocumentUpdateResult> UpdateBookmrkValue(string documentId, IEnumerable<DestinationPoint> destinationPoints, string value);
        Task<DocumentCheckResult> GetDocumentUrlByID(DocumentCheckResult result);
    }
}
