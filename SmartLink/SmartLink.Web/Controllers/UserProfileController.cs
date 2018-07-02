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
    public class UserProfileController : ApiController
    {
        protected readonly IUserProfileService _userProfileService;
        public UserProfileController(IUserProfileService userProfileService)
        {
            _userProfileService = userProfileService;
        }

        [HttpGet]
        [Route("api/UserProfile")]
        public IHttpActionResult GetUserProfile()
        {
            var retValue = _userProfileService.GetCurrentUser();
            return Ok(retValue);
        }
    }
}