// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using AutoMapper;
using Microsoft.Azure;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using SmartLink.Entity;
using SmartLink.Service;
using SmartLink.Web.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace SmartLink.Web.Controllers
{
	[APIAuthorize]
    public class DestinationPointController : ApiController
    {
        protected readonly IDestinationService _destinationService;
        protected readonly IMapper _mapper;
		private readonly string clientId;
		private readonly string aadInstance;
		private readonly string tenantId;
		private readonly string appKey;
		private readonly string resourceId;
		private readonly string sharePointUrl;
		private readonly string authority;

		public DestinationPointController(IDestinationService destinationService, IMapper mapper)
		{
			_destinationService = destinationService;
			_mapper = mapper;
			clientId = CloudConfigurationManager.GetSetting("ida:ClientId");
			aadInstance = CloudConfigurationManager.GetSetting("ida:AADInstance");
			tenantId = CloudConfigurationManager.GetSetting("ida:TenantId");
			appKey = CloudConfigurationManager.GetSetting("ida:ClientSecret");
			resourceId = CloudConfigurationManager.GetSetting("ResourceId");
			sharePointUrl = CloudConfigurationManager.GetSetting("SharePointUrl");
			authority = aadInstance + tenantId;
		}

        [HttpPost]
        [Route("api/DestinationPoint")]
        public async Task<IHttpActionResult> Post([FromBody]DestinationPointForm destinationPointAdded)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest("Invalid posted data.");
            }

            try
            {
                var destinationPoint = _mapper.Map<DestinationPoint>(destinationPointAdded);
                var catalogName = HttpUtility.UrlDecode(destinationPointAdded.CatalogName);
                var documentId = HttpUtility.UrlDecode(destinationPointAdded.DocumentId);
                return Ok(await _destinationService.AddDestinationPoint(catalogName, documentId, destinationPoint));
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }

        [HttpDelete]
        [Route("api/DestinationPoint")]
        public async Task<IHttpActionResult> DeleteSourcePoint(string id)
        {
            await _destinationService.DeleteDestinationPoint(Guid.Parse(id));
            return Ok();
        }

        [HttpPost]
        [Route("api/DeleteSelectedDestinationPoint")]
        public async Task<IHttpActionResult> DeleteSelectedDestinationPoint([FromBody]IEnumerable<Guid> seletedIds)
        {
            await _destinationService.DeleteSelectedDestinationPoint(seletedIds);
            return Ok();
        }

        [HttpGet]
        [Route("api/DestinationPointCatalog")]
        public async Task<IHttpActionResult> GetDestinationPointCatalog(string fileName, string documentId)
        {
            var retValue = await _destinationService.GetDestinationCatalog(HttpUtility.UrlDecode(fileName), HttpUtility.UrlDecode(documentId));
            return Ok(retValue);
        }

        [HttpGet]
        [Route("api/DestinationPoint")]
        public async Task<IHttpActionResult> GetDestinationPointBySourcePoint(string sourcePointId)
        {
            var retValue = await _destinationService.GetDestinationPointBySourcePoint(Guid.Parse(sourcePointId));
            return Ok(retValue);
        }

		[HttpGet]
		[Route("api/GraphAccessToken")]
		public async Task<IHttpActionResult> GetGraphAccessToken()
		{
			ClientCredential clientCred = new ClientCredential(clientId, appKey);
			var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as System.IdentityModel.Tokens.BootstrapContext;
			string userName = ClaimsPrincipal.Current.FindFirst(ClaimTypes.Upn) != null ? ClaimsPrincipal.Current.FindFirst(ClaimTypes.Upn).Value : ClaimsPrincipal.Current.FindFirst(ClaimTypes.Email).Value;
			string userAccessToken = bootstrapContext.Token;
			UserAssertion userAssertion = new UserAssertion(userAccessToken, "urn:ietf:params:oauth:grant-type:jwt-bearer", userName);

			string userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
			AuthenticationContext authContext = new AuthenticationContext(authority);

			var result = await authContext.AcquireTokenAsync(resourceId, clientCred, userAssertion);
			return Ok(result.AccessToken);
		}

        [HttpGet]
        [Route("api/CustomFormats")]
        public async Task<IHttpActionResult> GetCustomFormats()
        {
            var retValue = await _destinationService.GetCustomFormats();
            return Ok(retValue);
        }

        [HttpPut]
        [Route("api/UpdateDestinationPointCustomFormat")]
        public async Task<IHttpActionResult> UpdateDestinationPointCustomFormat([FromBody]DestinationPointForm destinationPointAdded)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest("Invalid posted data.");
            }

            try
            {
                var destinationPoint = _mapper.Map<DestinationPoint>(destinationPointAdded);
                return Ok(await _destinationService.UpdateDestinationPointCustomFormat(destinationPoint));
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }

        [HttpGet]
        [Route("api/SharePointAccessToken")]
        public async Task<IHttpActionResult> GetSharePointAccessToken()
        {
			ClientCredential clientCred = new ClientCredential(clientId, appKey);
			var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as System.IdentityModel.Tokens.BootstrapContext;
			string userName = ClaimsPrincipal.Current.FindFirst(ClaimTypes.Upn) != null ? ClaimsPrincipal.Current.FindFirst(ClaimTypes.Upn).Value : ClaimsPrincipal.Current.FindFirst(ClaimTypes.Email).Value;
			string userAccessToken = bootstrapContext.Token;
			UserAssertion userAssertion = new UserAssertion(userAccessToken, "urn:ietf:params:oauth:grant-type:jwt-bearer", userName);

			string userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
			AuthenticationContext authContext = new AuthenticationContext(authority);

			var result = await authContext.AcquireTokenAsync(sharePointUrl, clientCred, userAssertion);
			return Ok(result.AccessToken);
        }
    }
}
