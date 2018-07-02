// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using ApplicationCore.Artifacts;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json.Linq;

namespace ApplicationCore.Interfaces
{
    public interface IDocumentRepository : IArtifactFactory<Document>
    {
        Task<JObject> UploadDocumentAsync(string siteId, string folder, IFormFile file, string requestId = "");

        Task<JObject> UploadDocumentTeamAsync(string opportunityName, string docType, IFormFile file, string requestId = "");
    }
}