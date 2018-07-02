// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Security.Claims;
using System.Text;
using Microsoft.AspNetCore.Http;
using ApplicationCore.Interfaces;

namespace Infrastructure.Identity
{
    /// <summary>
    /// Adapter to pass the ClaimsPrincipal from the HTTPContext in ASP.Net Core
    /// </summary>
    public class UserIdentityContext : IUserContext
    {
        private readonly IHttpContextAccessor _accessor;

        public UserIdentityContext(IHttpContextAccessor accessor)
        {
            _accessor = accessor;
        }

        public ClaimsPrincipal User => _accessor.HttpContext.User;
    }
}
