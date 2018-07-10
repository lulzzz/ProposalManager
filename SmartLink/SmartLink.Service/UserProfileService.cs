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
using System.Security.Principal;
using System.Security.Claims;
using System.Threading;

namespace SmartLink.Service
{
    public class UserProfileService : IUserProfileService
    {
        public UserProfile GetCurrentUser()
        {
            var currentUserProfile = new UserProfile();
            var icp = Thread.CurrentPrincipal as ClaimsPrincipal;
            if (icp != null)
            {
                currentUserProfile.Email = icp.Identity.Name;
                var nameClaim = icp.Claims.FirstOrDefault(o => o.Type == "name");
                if (nameClaim != null)
                {
                    currentUserProfile.Username = nameClaim.Value;
                }
            }
            return currentUserProfile;
        }
    }
}
