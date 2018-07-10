// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using AutoMapper;
using SmartLink.Entity;
using SmartLink.Service;
using SmartLink.Web.ViewModel;
using System;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace SmartLink.Web.Controllers
{
    public class RecentFileController : ApiController
    {
        protected readonly IRecentFileService _recentFileService;
        protected readonly IMapper _mapper;
        public RecentFileController(IRecentFileService recentFileService, IMapper mapper)
        {
            _recentFileService = recentFileService;
            _mapper = mapper;
        }

        [HttpGet]
        [Route("api/RecentFiles")]
        public async Task<IHttpActionResult> GetRecentFiles()
        {
            var retValue = await _recentFileService.GetRecentFiles();
            return Ok(retValue);
        }

        [HttpPost]
        [Route("api/RecentFile")]
        public async Task<IHttpActionResult> Post([FromBody]CatalogViewModel catalogAdded)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest("Invalid posted data.");
            }

            try
            {
                var catalogName = HttpUtility.UrlDecode(catalogAdded.Name);
                var documentId = HttpUtility.UrlDecode(catalogAdded.DocumentId);
                return Ok(await _recentFileService.AddRecentFile(new SourceCatalog() { Name = catalogName, DocumentId = documentId }));
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }
    }
}