// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using AutoMapper;
using SmartLink.Entity;
using SmartLink.Service;
using SmartLink.Web.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Security.Principal;
using System.Security.Claims;
using System.Web;
using Newtonsoft.Json;

namespace SmartLink.Web.Controllers
{
    [APIAuthorize]
    public class SourcePointController : ApiController
    {
        protected readonly ISourceService _sourceService;
        protected readonly IMapper _mapper;
        public SourcePointController(ISourceService sourceService, IMapper mapper)
        {
            _sourceService = sourceService;
            _mapper = mapper;
        }

        [HttpPost]
        [Route("api/SourcePoint")]
        public async Task<IHttpActionResult> Post([FromBody]SourcePointForm sourcePointAdded)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest("Invalid posted data.");
            }

            try
            {
                var sourcePoint = _mapper.Map<SourcePoint>(sourcePointAdded);
                var catalogName = HttpUtility.UrlDecode(sourcePointAdded.CatalogName);
                var documentId = HttpUtility.UrlDecode(sourcePointAdded.DocumentId);
                return Ok(await _sourceService.AddSourcePoint(catalogName, documentId, sourcePoint));
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }

        [HttpGet]
        [Route("api/SourcePointCatalog")]
        public async Task<IHttpActionResult> GetSourcePointCatalog(string fileName, string documentId)
        {
            var retValue = await _sourceService.GetSourceCatalog(HttpUtility.UrlDecode(fileName), HttpUtility.UrlDecode(documentId));
            return Ok(retValue);
        }

        [HttpGet]
        [Route("api/SourcePointCatalog")]
        public async Task<IHttpActionResult> GetSourcePointCatalog(string documentId)
        {
            var retValue = await _sourceService.GetSourceCatalog(HttpUtility.UrlDecode(documentId));
            return Ok(retValue);
        }

        [HttpPost]
        [Route("api/PublishSourcePoints")]
        public async Task<IHttpActionResult> PublishSourcePoints([FromBody]IEnumerable<PublishSourcePointForm> sourcePointPublishForm)
        {
            if (!ModelState.IsValid || sourcePointPublishForm.Count() == 0)
            {
                return BadRequest("Invalid posted data.");
            }

            var retValue = await _sourceService.PublishSourcePointList(sourcePointPublishForm);

            return Ok(retValue);
        }

        [HttpPut]
        [Route("api/SourcePoint")]
        public async Task<IHttpActionResult> EditSourcePoint([FromBody]SourcePointForm sourcePointAdded)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest("Invalid posted data.");
            }

            try
            {
                var sourcePoint = _mapper.Map<SourcePoint>(sourcePointAdded);

                return Ok(await _sourceService.EditSourcePoint(sourcePoint));
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }

        [HttpDelete]
        [Route("api/SourcePoint")]
        public async Task<IHttpActionResult> DeleteSourcePoint(string id)
        {
            var retValue = await _sourceService.DeleteSourcePoint(new Guid(id));
            return Ok();
        }

        [HttpPost]
        [Route("api/DeleteSelectedSourcePoint")]
        public async Task<IHttpActionResult> DeleteSelectedSourcePoint([FromBody]IEnumerable<Guid> seletedIds)
        {
            await _sourceService.DeleteSelectedSourcePoint(seletedIds);
            return Ok();
        }

        [HttpPost]
        [Route("api/CloneCheckFile")]
        public async Task<IHttpActionResult> CloneCheckFile([FromBody]IEnumerable<CloneForm> files)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest("Invalid posted data.");
            }

            try
            {
                return Ok(await _sourceService.CheckCloneFileStatus(files));
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }

        [HttpPost]
        [Route("api/CloneFiles")]
        public async Task<IHttpActionResult> CloneFiles([FromBody]IEnumerable<CloneForm> files)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest("Invalid posted data.");
            }

            try
            {
                foreach (var item in files)
                {
                    item.DestinationFileUrl = HttpUtility.UrlDecode(item.DestinationFileUrl);
                }
                await _sourceService.CloneFiles(files);
                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }
    }
}